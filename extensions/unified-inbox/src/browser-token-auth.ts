// ============================================================================
// CDP-based token extraction from Microsoft web apps
// Extracts Graph API tokens from MSAL sessionStorage in live browser tabs.
// Each instance manages ONE tab (e.g. Outlook or Teams).
//
// Includes automatic token recovery when tokens expire:
//   Phase 1: Clear stale MSAL cache → reload browser tab → poll for fresh token
//   Phase 2: Use refresh token + Azure AD /oauth2/v2.0/token to acquire new token
//   Phase 3: Intercept Bearer token from live page requests via CDP Fetch domain
//            (handles MSAL v4 encrypted cache where localStorage entries are AES-GCM encrypted)
// ============================================================================

import http from "node:http";
import type { IMsAuthProvider } from "./types.js";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

export type BrowserTokenAuthOptions = {
  cdpPort: number;   // e.g. 18801
  pageUrl: string;   // e.g. "https://outlook.office.com" or "https://teams.microsoft.com"
  pageName: string;  // e.g. "Outlook" or "Teams" (for logging)
  /** Token resource to extract. Default: "graph.microsoft.com".
   *  Use "outlook.office.com" for the Outlook REST API token (has Mail.Send). */
  resource?: string;
};

type CdpTarget = {
  type: string;
  url: string;
  webSocketDebuggerUrl: string;
};

type ExtractedToken = {
  token: string;
  expiresOn: number;
  target: string;
};

// Re-extract from sessionStorage 5 min before we consider it expired
const REFRESH_BUFFER_MS = 5 * 60 * 1000;
// Auto re-extraction interval: 10 minutes
const AUTO_REFRESH_INTERVAL_MS = 10 * 60 * 1000;
// Wait for page readiness: poll every 3s for up to 60s
const PAGE_READY_POLL_MS = 3_000;
const PAGE_READY_TIMEOUT_MS = 60_000;
// Wait for CDP port: poll every 5s for up to 150s (browser may start well after daemon)
const CDP_CONNECT_POLL_MS = 5_000;
const CDP_CONNECT_TIMEOUT_MS = 150_000;
// Recovery: poll every 3s for up to 30s after clearing stale tokens + reload
const RECOVERY_POLL_MS = 3_000;
const RECOVERY_POLL_TIMEOUT_MS = 30_000;
// Fetch intercept: wait up to 15s for a Bearer token from live page requests
const FETCH_INTERCEPT_TIMEOUT_MS = 15_000;

// ---------------------------------------------------------------------------
// JS builders — generate code to evaluate in browser context via CDP
// ---------------------------------------------------------------------------

/** Build JS to extract a token from MSAL cache (localStorage + sessionStorage). */
function buildExtractTokenJs(resource: string): string {
  return `(() => {
  const stores = [localStorage, sessionStorage];
  for (const store of stores) {
    for (const key of Object.keys(store)) {
      if (!key.toLowerCase().includes('accesstoken')) continue;
      try {
        const entry = JSON.parse(store.getItem(key));
        if (!entry || !entry.secret) continue;
        if (key.toLowerCase().includes('${resource.toLowerCase()}')) {
          return { token: entry.secret, expiresOn: Number(entry.expires_on || entry.expiresOn) || 0, target: entry.target || '' };
        }
      } catch {}
    }
  }
  return null;
})()`;
}

/** Build JS to clear expired/stale accesstoken entries for a resource from localStorage. */
function buildClearStaleTokensJs(resource: string): string {
  return `(() => {
  let cleared = 0;
  for (const key of Object.keys(localStorage)) {
    if (!key.toLowerCase().includes('accesstoken')) continue;
    try {
      const entry = JSON.parse(localStorage.getItem(key));
      if (!entry || !entry.secret) continue;
      const expiresOn = Number(entry.expires_on || entry.expiresOn) || 0;
      const isExpired = expiresOn > 0 && expiresOn * 1000 < Date.now();
      if (isExpired || key.toLowerCase().includes('${resource.toLowerCase()}')) {
        localStorage.removeItem(key);
        cleared++;
      }
    } catch {}
  }
  return { cleared };
})()`;
}

/** Build JS to click on a mail/chat item and dispatch focus events to trigger API calls. */
function buildTriggerPageActivityJs(): string {
  return `(() => {
  // Try clicking on a mail/chat item to trigger a Graph API call
  const selectors = ['[data-convid]', '[role="option"]', '[role="listitem"] [role="row"]', '[role="row"]'];
  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (el) {
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
      return { clicked: sel };
    }
  }
  // Fallback: dispatch focus + visibilitychange to nudge background fetches
  window.dispatchEvent(new Event('focus'));
  document.dispatchEvent(new Event('visibilitychange'));
  return { clicked: null, fallback: true };
})()`;
}

/** Build JS to acquire a token via refresh token + Azure AD token endpoint.
 *  Runs entirely in browser context (uses page's fetch + localStorage). */
function buildAcquireTokenViaRefreshJs(resource: string): string {
  return `(async () => {
  // 1. Find refresh token in MSAL cache
  let refreshToken = null;
  for (const key of Object.keys(localStorage)) {
    if (key.toLowerCase().includes("refreshtoken")) {
      try {
        const entry = JSON.parse(localStorage.getItem(key));
        if (entry && entry.secret) { refreshToken = entry.secret; break; }
      } catch {}
    }
  }
  if (!refreshToken) return { error: "No refresh token found in localStorage" };

  // 2. Find MSAL account info (clientId, tenantId, homeAccountId)
  let clientId = null;
  for (const key of Object.keys(localStorage)) {
    if (key.startsWith("msal.") && key.includes("active-account")) {
      clientId = key.split(".")[1];
      break;
    }
  }
  if (!clientId) return { error: "No MSAL client ID found" };

  const accountKeysRaw = localStorage.getItem("msal.account.keys");
  if (!accountKeysRaw) return { error: "No MSAL account keys" };
  const accountKeys = JSON.parse(accountKeysRaw);
  if (!accountKeys.length) return { error: "Empty MSAL accounts list" };
  const accountData = JSON.parse(localStorage.getItem(accountKeys[0]));
  if (!accountData) return { error: "MSAL account data not found" };

  const tenantId = accountData.realm;
  const homeAccountId = accountData.home_account_id;
  if (!tenantId || !homeAccountId) return { error: "Missing tenantId or homeAccountId" };

  // 3. Exchange refresh token for access token via Azure AD
  const scope = "https://${resource}/.default";
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: scope,
  });

  const resp = await fetch(
    "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!resp.ok) {
    const text = await resp.text();
    return { error: "Azure AD " + resp.status + ": " + text.slice(0, 300) };
  }

  const data = await resp.json();
  const expiresOn = Math.floor(Date.now() / 1000) + (data.expires_in || 3600);
  const target = data.scope || "";

  // 4. Store token in MSAL v2 cache format so extraction picks it up
  const tokenKey = homeAccountId + "-login.windows.net-accesstoken-" + clientId + "-" + tenantId + "-" + target.toLowerCase();
  localStorage.setItem(tokenKey, JSON.stringify({
    home_account_id: homeAccountId,
    environment: "login.windows.net",
    credential_type: "AccessToken",
    client_id: clientId,
    secret: data.access_token,
    realm: tenantId,
    target: target,
    cached_at: String(Math.floor(Date.now() / 1000)),
    expires_on: String(expiresOn),
    extended_expires_on: String(expiresOn + 3600),
    token_type: "Bearer",
  }));

  // 5. Update refresh token if rotated by Azure AD
  if (data.refresh_token) {
    for (const key of Object.keys(localStorage)) {
      if (key.toLowerCase().includes("refreshtoken")) {
        try {
          const entry = JSON.parse(localStorage.getItem(key));
          if (entry && entry.secret) {
            entry.secret = data.refresh_token;
            localStorage.setItem(key, JSON.stringify(entry));
          }
        } catch {}
      }
    }
  }

  return {
    success: true,
    token: data.access_token,
    expiresOn: expiresOn,
    scopes: target.slice(0, 200),
  };
})()`;
}

/** Check if a token is still fresh (expires more than 5 min from now) */
function isTokenFresh(expiresOn: number): boolean {
  return expiresOn * 1000 > Date.now() + REFRESH_BUFFER_MS;
}

/** Decode JWT payload (base64url → JSON) to extract exp and scp claims.
 *  No signature verification — we trust the browser's own tokens. */
function decodeJwtClaims(token: string): { exp: number; scp: string } | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    // base64url → base64 → decode
    const payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = Buffer.from(payload, "base64").toString("utf-8");
    const claims = JSON.parse(json) as { exp?: number; scp?: string };
    return {
      exp: claims.exp ?? 0,
      scp: claims.scp ?? "",
    };
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// BrowserTokenAuth
// ---------------------------------------------------------------------------

/**
 * Extracts Graph API tokens directly from MSAL token cache in a live
 * Microsoft web app tab (Outlook, Teams, etc.) via Chrome DevTools Protocol.
 *
 * No OAuth dialogs, no client IDs, no admin consent needed — uses tokens
 * the web app has already acquired for the logged-in user.
 *
 * When tokens expire, automatic recovery kicks in:
 *   Phase 1: Clear stale MSAL entries → reload tab → poll for fresh token
 *   Phase 2: Use refresh token to call Azure AD directly → store result in MSAL format
 *   Phase 3: Intercept Bearer token from live page requests via CDP Fetch domain
 */
export class BrowserTokenAuth implements IMsAuthProvider {
  private cachedToken: string | null = null;
  private expiresAt = 0;
  private refreshTimer: ReturnType<typeof setInterval> | null = null;
  private readonly cdpPort: number;
  private readonly pageUrl: string;
  private readonly pageUrlHost: string;
  private readonly pageName: string;
  private readonly resource: string;
  private readonly extractTokenJs: string;
  private readonly log: Logger;

  // Recovery state
  private hasEverSucceeded = false;
  private isRecovering = false;
  private onTokenRecovered: (() => void) | null = null;

  constructor(opts: BrowserTokenAuthOptions, log: Logger) {
    this.cdpPort = opts.cdpPort;
    this.pageUrl = opts.pageUrl;
    this.pageUrlHost = new URL(opts.pageUrl).hostname;
    this.pageName = opts.pageName;
    this.resource = opts.resource ?? "graph.microsoft.com";
    this.extractTokenJs = buildExtractTokenJs(this.resource);
    this.log = log;
  }

  // ---------------------------------------------------------------------------
  // IMsAuthProvider implementation
  // ---------------------------------------------------------------------------

  isAuthenticated(): boolean {
    return this.cachedToken !== null;
  }

  async getAccessToken(): Promise<string> {
    // Re-extract if near expiry
    if (this.cachedToken && Date.now() >= this.expiresAt - REFRESH_BUFFER_MS) {
      this.log.info(`unified-inbox: ${this.pageName}: token near expiry, re-extracting...`);
      await this.extractTokenFromTab();
    }

    if (!this.cachedToken) {
      throw new Error(
        `${this.pageName} token not available. Is the browser running and logged into Microsoft 365?`,
      );
    }

    return this.cachedToken;
  }

  async loadPersistedTokens(): Promise<boolean> {
    try {
      await this.extractTokenFromTab();
      return this.cachedToken !== null;
    } catch (err) {
      this.log.warn(
        `unified-inbox: ${this.pageName}: failed to extract token: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  startAutoRefresh(onError: (error: string) => Promise<void>): void {
    this.stopAutoRefresh();
    this.refreshTimer = setInterval(async () => {
      try {
        await this.extractTokenFromTab();
        if (!this.cachedToken) {
          await onError(
            `${this.pageName}: no token found in sessionStorage. Is the user logged in?`,
          );
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        this.log.error(`unified-inbox: ${this.pageName}: auto-refresh failed: ${msg}`);
        await onError(`${this.pageName} token refresh failed: ${msg}`);
      }
    }, AUTO_REFRESH_INTERVAL_MS);
  }

  stopAutoRefresh(): void {
    if (this.refreshTimer) {
      clearInterval(this.refreshTimer);
      this.refreshTimer = null;
    }
  }

  /** Register callback invoked when token recovers after failure */
  setOnTokenRecovered(cb: () => void): void {
    this.onTokenRecovered = cb;
  }

  /** Force token refresh/recovery — for remote Telegram commands.
   *  Returns true if a fresh token was obtained. */
  async forceRefresh(): Promise<boolean> {
    this.log.info(`unified-inbox: ${this.pageName}: force refresh requested`);
    try {
      const wsUrl = await this.findOrCreateTab();

      // Try direct extraction first
      const immediate = await this.evaluateTokenExtraction(wsUrl);
      if (immediate && isTokenFresh(immediate.expiresOn)) {
        this.applyToken(immediate);
        return true;
      }

      // Token missing or stale — run full recovery
      const result = await this.attemptRecovery(wsUrl);
      if (result.recovered) {
        this.log.info(
          `unified-inbox: ${this.pageName}: force refresh succeeded via ${result.method}`,
        );
        this.onTokenRecovered?.();
        return true;
      }

      this.log.warn(
        `unified-inbox: ${this.pageName}: force refresh failed — no fresh token recovered`,
      );
      return false;
    } catch (err) {
      this.log.error(
        `unified-inbox: ${this.pageName}: force refresh error: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  // ---------------------------------------------------------------------------
  // CDP communication — token extraction
  // ---------------------------------------------------------------------------

  /** Find an existing tab or create a new one, then extract the MSAL token.
   *  If token is stale and we've succeeded before, attempt automatic recovery. */
  private async extractTokenFromTab(): Promise<void> {
    const wsUrl = await this.findOrCreateTab();

    // Try immediate extraction first (tab may already be loaded)
    const immediate = await this.evaluateTokenExtraction(wsUrl);
    if (immediate) {
      if (isTokenFresh(immediate.expiresOn)) {
        this.applyToken(immediate);
        return;
      }

      // Token found but expired — attempt recovery if we've ever had a good token
      if (this.hasEverSucceeded && !this.isRecovering) {
        this.log.warn(
          `unified-inbox: ${this.pageName}: extracted token is expired, attempting recovery...`,
        );
        const result = await this.attemptRecovery(wsUrl);
        if (result.recovered) {
          this.log.info(
            `unified-inbox: ${this.pageName}: token recovered via ${result.method}`,
          );
          this.onTokenRecovered?.();
          return;
        }
      }

      // Apply stale token as fallback (better than null for transient issues)
      this.applyToken(immediate);
      return;
    }

    // No token in localStorage/sessionStorage — try Fetch intercept before
    // entering the slow polling loop. This handles MSAL v4 encrypted cache
    // where the page is already loaded and making authenticated API calls.
    this.log.info(
      `unified-inbox: ${this.pageName}: no token in storage, trying request intercept...`,
    );
    const intercepted = await this.extractTokenViaFetchIntercept(wsUrl);
    if (intercepted) {
      if (isTokenFresh(intercepted.expiresOn)) {
        this.applyToken(intercepted);
        return;
      }
      // Intercepted but expired — apply as fallback
      this.applyToken(intercepted);
      return;
    }

    // No token found — page may need time to load, poll until ready
    this.log.info(`unified-inbox: ${this.pageName}: waiting for page to load...`);
    const deadline = Date.now() + PAGE_READY_TIMEOUT_MS;

    while (Date.now() < deadline) {
      await sleep(PAGE_READY_POLL_MS);
      const result = await this.evaluateTokenExtraction(wsUrl);
      if (result) {
        if (isTokenFresh(result.expiresOn)) {
          this.applyToken(result);
          return;
        }
        // Expired token found during polling — try recovery
        if (this.hasEverSucceeded && !this.isRecovering) {
          this.log.warn(
            `unified-inbox: ${this.pageName}: polled token is expired, attempting recovery...`,
          );
          const recovery = await this.attemptRecovery(wsUrl);
          if (recovery.recovered) {
            this.log.info(
              `unified-inbox: ${this.pageName}: token recovered via ${recovery.method}`,
            );
            this.onTokenRecovered?.();
            return;
          }
        }
        this.applyToken(result);
        return;
      }
    }

    // Timed out with no token — attempt recovery if we've succeeded before
    if (this.hasEverSucceeded && !this.isRecovering) {
      this.log.warn(
        `unified-inbox: ${this.pageName}: no token after ${PAGE_READY_TIMEOUT_MS / 1000}s, attempting recovery...`,
      );
      const result = await this.attemptRecovery(wsUrl);
      if (result.recovered) {
        this.log.info(
          `unified-inbox: ${this.pageName}: token recovered via ${result.method}`,
        );
        this.onTokenRecovered?.();
        return;
      }
    }

    this.log.warn(
      `unified-inbox: ${this.pageName}: no MSAL token found after ${PAGE_READY_TIMEOUT_MS / 1000}s`,
    );
    this.cachedToken = null;
  }

  private applyToken(extracted: ExtractedToken): void {
    this.cachedToken = extracted.token;
    // MSAL stores expiresOn as Unix epoch seconds
    this.expiresAt = extracted.expiresOn * 1000;
    if (isTokenFresh(extracted.expiresOn)) {
      this.hasEverSucceeded = true;
    }
    this.log.info(
      `unified-inbox: ${this.pageName}: token extracted (expires ${new Date(this.expiresAt).toLocaleTimeString()}, scopes: ${extracted.target})`,
    );
  }

  // ---------------------------------------------------------------------------
  // Token recovery
  // ---------------------------------------------------------------------------

  /** Full recovery pipeline:
   *  Phase 1: Clear stale tokens → reload tab → poll 30s for fresh token
   *  Phase 2: Use refresh token via Azure AD to acquire a new token directly
   *  Phase 3: Intercept Bearer token from live page requests via CDP Fetch domain */
  private async attemptRecovery(
    wsUrl: string,
  ): Promise<{ recovered: boolean; method?: string }> {
    if (this.isRecovering) return { recovered: false };
    this.isRecovering = true;

    try {
      // --- Phase 1: Clear stale MSAL entries + reload browser tab ---
      this.log.info(
        `unified-inbox: ${this.pageName}: recovery phase 1 — clearing stale tokens + reloading tab`,
      );

      const clearResult = await this.evaluateJs<{ cleared: number }>(
        wsUrl,
        buildClearStaleTokensJs(this.resource),
      );
      this.log.info(
        `unified-inbox: ${this.pageName}: cleared ${clearResult?.cleared ?? 0} stale token entries`,
      );

      await this.reloadTab(wsUrl);

      // Poll for fresh token after reload
      const phase1Deadline = Date.now() + RECOVERY_POLL_TIMEOUT_MS;
      while (Date.now() < phase1Deadline) {
        await sleep(RECOVERY_POLL_MS);
        const result = await this.evaluateTokenExtraction(wsUrl);
        if (result && isTokenFresh(result.expiresOn)) {
          this.applyToken(result);
          return { recovered: true, method: "tab-reload" };
        }
      }

      this.log.info(
        `unified-inbox: ${this.pageName}: phase 1 did not produce a fresh token`,
      );

      // --- Phase 2: Acquire token via Azure AD refresh token exchange ---
      this.log.info(
        `unified-inbox: ${this.pageName}: recovery phase 2 — acquiring token via Azure AD refresh token`,
      );

      const acquireResult = await this.evaluateJs<{
        success?: boolean;
        error?: string;
        token?: string;
        expiresOn?: number;
        scopes?: string;
      }>(wsUrl, buildAcquireTokenViaRefreshJs(this.resource), true);

      if (acquireResult?.success && acquireResult.token && acquireResult.expiresOn) {
        this.applyToken({
          token: acquireResult.token,
          expiresOn: acquireResult.expiresOn,
          target: acquireResult.scopes ?? "",
        });
        return { recovered: true, method: "refresh-token-exchange" };
      }

      if (acquireResult?.error) {
        this.log.error(
          `unified-inbox: ${this.pageName}: refresh token exchange failed: ${acquireResult.error}`,
        );
      }

      // Final check — the acquire step stores the token in localStorage,
      // so extraction should find it even if the return value was lost
      const finalCheck = await this.evaluateTokenExtraction(wsUrl);
      if (finalCheck && isTokenFresh(finalCheck.expiresOn)) {
        this.applyToken(finalCheck);
        return { recovered: true, method: "refresh-token-exchange" };
      }

      // --- Phase 3: Intercept Bearer token from live page requests ---
      // When MSAL v4 encrypts the cache, the page is still functional
      // and making authenticated API calls — we just intercept one.
      this.log.info(
        `unified-inbox: ${this.pageName}: recovery phase 3 — intercepting Bearer token via CDP Fetch domain`,
      );
      const intercepted = await this.extractTokenViaFetchIntercept(wsUrl);
      if (intercepted && isTokenFresh(intercepted.expiresOn)) {
        this.applyToken(intercepted);
        return { recovered: true, method: "request-intercept" };
      }

      return { recovered: false };
    } finally {
      this.isRecovering = false;
    }
  }

  /** Reload the browser tab via CDP Page.reload */
  private reloadTab(wsUrl: string): Promise<void> {
    return new Promise((resolve) => {
      const ws = new WebSocket(wsUrl);
      let settled = false;

      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          ws.close();
          resolve();
        }
      }, 10_000);

      ws.addEventListener("open", () => {
        ws.send(JSON.stringify({ id: 1, method: "Page.reload", params: {} }));
        // Give the page a moment to start reloading, then resolve
        setTimeout(() => {
          if (!settled) {
            settled = true;
            clearTimeout(timer);
            ws.close();
            resolve();
          }
        }, 2_000);
      });

      ws.addEventListener("error", () => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve();
        }
      });
    });
  }

  /** Generic CDP evaluate — runs JS in page context and returns typed result. */
  private evaluateJs<T>(
    wsUrl: string,
    expression: string,
    awaitPromise = false,
  ): Promise<T | null> {
    return new Promise((resolve) => {
      const ws = new WebSocket(wsUrl);
      let settled = false;

      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          ws.close();
          resolve(null);
        }
      }, 15_000);

      ws.addEventListener("open", () => {
        ws.send(
          JSON.stringify({
            id: 1,
            method: "Runtime.evaluate",
            params: { expression, returnByValue: true, awaitPromise },
          }),
        );
      });

      ws.addEventListener("message", (event: MessageEvent) => {
        if (settled) return;
        try {
          const msg = JSON.parse(String(event.data)) as {
            id?: number;
            result?: { result?: { value?: T | null } };
          };
          if (msg.id === 1) {
            settled = true;
            clearTimeout(timer);
            ws.close();
            resolve(msg.result?.result?.value ?? null);
          }
        } catch {
          // ignore parse errors from other messages
        }
      });

      ws.addEventListener("error", () => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve(null);
        }
      });
    });
  }

  // ---------------------------------------------------------------------------
  // CDP tab management
  // ---------------------------------------------------------------------------

  /** List CDP targets, find one matching our pageUrlHost, or create a new tab.
   *  Retries CDP connection for up to 90s if the browser isn't running yet. */
  private async findOrCreateTab(): Promise<string> {
    const deadline = Date.now() + CDP_CONNECT_TIMEOUT_MS;
    let lastError: Error | null = null;

    while (Date.now() < deadline) {
      try {
        const targets = await this.listTargets();

        // Look for existing tab
        const existing = targets.find(
          (t) => t.type === "page" && t.url.includes(this.pageUrlHost),
        );
        if (existing) {
          return existing.webSocketDebuggerUrl;
        }

        // No matching tab — create one
        this.log.info(`unified-inbox: ${this.pageName}: no tab found, creating...`);
        const created = await this.createTab(this.pageUrl);
        return created.webSocketDebuggerUrl;
      } catch (err) {
        lastError = err instanceof Error ? err : new Error(String(err));
        // Only retry on connection errors (browser not running yet)
        if (!lastError.message.includes("ECONNREFUSED")) {
          throw lastError;
        }
        this.log.info(
          `unified-inbox: ${this.pageName}: browser not ready, retrying in ${CDP_CONNECT_POLL_MS / 1000}s...`,
        );
        await sleep(CDP_CONNECT_POLL_MS);
      }
    }

    throw lastError ?? new Error("CDP connection timed out");
  }

  /** GET http://127.0.0.1:{cdpPort}/json — list all CDP targets */
  private listTargets(): Promise<CdpTarget[]> {
    return new Promise((resolve, reject) => {
      http
        .get(`http://127.0.0.1:${this.cdpPort}/json`, (res) => {
          let data = "";
          res.on("data", (chunk: Buffer) => (data += chunk));
          res.on("end", () => {
            try {
              resolve(JSON.parse(data) as CdpTarget[]);
            } catch {
              reject(new Error("Failed to parse CDP targets"));
            }
          });
        })
        .on("error", (err) =>
          reject(
            new Error(
              `CDP connection failed (port ${this.cdpPort}): ${err.message}. Is the browser running?`,
            ),
          ),
        );
    });
  }

  /** PUT http://127.0.0.1:{cdpPort}/json/new?{url} — open a new tab */
  private createTab(url: string): Promise<CdpTarget> {
    return new Promise((resolve, reject) => {
      const req = http.request(
        `http://127.0.0.1:${this.cdpPort}/json/new?${encodeURIComponent(url)}`,
        { method: "PUT" },
        (res) => {
          let data = "";
          res.on("data", (chunk: Buffer) => (data += chunk));
          res.on("end", () => {
            try {
              resolve(JSON.parse(data) as CdpTarget);
            } catch {
              reject(new Error(`Failed to parse new tab response: ${data}`));
            }
          });
        },
      );
      req.on("error", (err) =>
        reject(new Error(`Failed to create tab: ${err.message}`)),
      );
      req.end();
    });
  }

  /** Open a WebSocket to the target, run Runtime.evaluate, parse the result */
  private evaluateTokenExtraction(wsUrl: string): Promise<ExtractedToken | null> {
    return new Promise((resolve) => {
      const ws = new WebSocket(wsUrl);
      let settled = false;

      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          ws.close();
          resolve(null);
        }
      }, 10_000);

      ws.addEventListener("open", () => {
        ws.send(
          JSON.stringify({
            id: 1,
            method: "Runtime.evaluate",
            params: {
              expression: this.extractTokenJs,
              returnByValue: true,
              awaitPromise: false,
            },
          }),
        );
      });

      ws.addEventListener("message", (event: MessageEvent) => {
        if (settled) return;
        try {
          const msg = JSON.parse(String(event.data)) as {
            id?: number;
            result?: { result?: { value?: ExtractedToken | null } };
          };
          if (msg.id === 1) {
            settled = true;
            clearTimeout(timer);
            ws.close();
            resolve(msg.result?.result?.value ?? null);
          }
        } catch {
          // ignore parse errors from other messages
        }
      });

      ws.addEventListener("error", () => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve(null);
        }
      });
    });
  }

  /** Intercept outgoing HTTP requests via CDP Fetch domain to capture Bearer tokens.
   *  Used when MSAL v4 encrypts the localStorage cache but the page is functional. */
  private extractTokenViaFetchIntercept(wsUrl: string): Promise<ExtractedToken | null> {
    return new Promise((resolve) => {
      const ws = new WebSocket(wsUrl);
      let settled = false;
      let msgId = 1;

      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          // Best-effort: disable Fetch domain before closing
          ws.send(JSON.stringify({ id: msgId++, method: "Fetch.disable", params: {} }));
          setTimeout(() => ws.close(), 500);
          this.log.warn(
            `unified-inbox: ${this.pageName}: fetch intercept timed out after ${FETCH_INTERCEPT_TIMEOUT_MS / 1000}s`,
          );
          resolve(null);
        }
      }, FETCH_INTERCEPT_TIMEOUT_MS);

      const finish = (token: ExtractedToken | null) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        // Disable Fetch domain before closing
        ws.send(JSON.stringify({ id: msgId++, method: "Fetch.disable", params: {} }));
        setTimeout(() => ws.close(), 500);
        resolve(token);
      };

      ws.addEventListener("open", () => {
        // Enable Fetch intercept for Graph API requests
        const enableId = msgId++;
        ws.send(
          JSON.stringify({
            id: enableId,
            method: "Fetch.enable",
            params: {
              patterns: [
                { urlPattern: `*${this.resource}*`, requestStage: "Request" },
              ],
            },
          }),
        );

        // After Fetch is enabled, inject JS to trigger page activity
        setTimeout(() => {
          if (settled) return;
          const triggerId = msgId++;
          ws.send(
            JSON.stringify({
              id: triggerId,
              method: "Runtime.evaluate",
              params: {
                expression: buildTriggerPageActivityJs(),
                returnByValue: true,
                awaitPromise: false,
              },
            }),
          );
        }, 500);
      });

      ws.addEventListener("message", (event: MessageEvent) => {
        if (settled) return;
        try {
          const msg = JSON.parse(String(event.data)) as {
            id?: number;
            method?: string;
            params?: {
              requestId?: string;
              request?: {
                url?: string;
                headers?: Record<string, string>;
              };
            };
          };

          // Handle Fetch.requestPaused events
          if (msg.method === "Fetch.requestPaused" && msg.params?.requestId) {
            const headers = msg.params.request?.headers ?? {};

            // Continue the request immediately so it doesn't hang
            ws.send(
              JSON.stringify({
                id: msgId++,
                method: "Fetch.continueRequest",
                params: { requestId: msg.params.requestId },
              }),
            );

            // Check for Authorization: Bearer header
            const authHeader =
              headers["Authorization"] ?? headers["authorization"] ?? "";
            if (authHeader.startsWith("Bearer ") && authHeader.length > 50) {
              const bearerToken = authHeader.slice(7);
              const claims = decodeJwtClaims(bearerToken);
              if (claims) {
                this.log.info(
                  `unified-inbox: ${this.pageName}: token extracted via request-intercept`,
                );
                finish({
                  token: bearerToken,
                  expiresOn: claims.exp,
                  target: claims.scp,
                });
                return;
              }
            }
          }
        } catch {
          // ignore parse errors
        }
      });

      ws.addEventListener("error", () => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve(null);
        }
      });
    });
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
