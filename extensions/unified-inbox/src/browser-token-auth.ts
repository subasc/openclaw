// ============================================================================
// CDP-based token extraction from Microsoft web apps
// Extracts Graph API tokens from MSAL sessionStorage in live browser tabs.
// Each instance manages ONE tab (e.g. Outlook or Teams).
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

/** JS executed inside the page context to extract the Graph API token from MSAL cache.
 *  MSAL v2 stores tokens in localStorage (not sessionStorage) on Outlook/Teams web. */
const EXTRACT_TOKEN_JS = `(() => {
  const stores = [localStorage, sessionStorage];
  for (const store of stores) {
    for (const key of Object.keys(store)) {
      if (!key.toLowerCase().includes('accesstoken')) continue;
      try {
        const entry = JSON.parse(store.getItem(key));
        if (!entry || !entry.secret) continue;
        if (key.toLowerCase().includes('graph.microsoft.com')) {
          return { token: entry.secret, expiresOn: Number(entry.expires_on || entry.expiresOn) || 0, target: entry.target || '' };
        }
      } catch {}
    }
  }
  return null;
})()`;

/**
 * Extracts Graph API tokens directly from MSAL token cache in a live
 * Microsoft web app tab (Outlook, Teams, etc.) via Chrome DevTools Protocol.
 *
 * No OAuth dialogs, no client IDs, no admin consent needed — uses tokens
 * the web app has already acquired for the logged-in user.
 */
export class BrowserTokenAuth implements IMsAuthProvider {
  private cachedToken: string | null = null;
  private expiresAt = 0;
  private refreshTimer: ReturnType<typeof setInterval> | null = null;
  private readonly cdpPort: number;
  private readonly pageUrl: string;
  private readonly pageUrlHost: string;
  private readonly pageName: string;
  private readonly log: Logger;

  constructor(opts: BrowserTokenAuthOptions, log: Logger) {
    this.cdpPort = opts.cdpPort;
    this.pageUrl = opts.pageUrl;
    this.pageUrlHost = new URL(opts.pageUrl).hostname;
    this.pageName = opts.pageName;
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

  // ---------------------------------------------------------------------------
  // CDP communication
  // ---------------------------------------------------------------------------

  /** Find an existing tab or create a new one, then extract the MSAL token */
  private async extractTokenFromTab(): Promise<void> {
    const wsUrl = await this.findOrCreateTab();

    // Try immediate extraction first (tab may already be loaded)
    const immediate = await this.evaluateTokenExtraction(wsUrl);
    if (immediate) {
      this.applyToken(immediate);
      return;
    }

    // Page may need time to load — poll until ready
    this.log.info(`unified-inbox: ${this.pageName}: waiting for page to load...`);
    const deadline = Date.now() + PAGE_READY_TIMEOUT_MS;

    while (Date.now() < deadline) {
      await sleep(PAGE_READY_POLL_MS);
      const result = await this.evaluateTokenExtraction(wsUrl);
      if (result) {
        this.applyToken(result);
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
    this.log.info(
      `unified-inbox: ${this.pageName}: token extracted (expires ${new Date(this.expiresAt).toLocaleTimeString()}, scopes: ${extracted.target})`,
    );
  }

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
              expression: EXTRACT_TOKEN_JS,
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
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
