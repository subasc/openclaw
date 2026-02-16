// ============================================================================
// Browser-based Microsoft auth: uses OAuth authorization code flow via
// Playwright browser with CDP to capture the auth code, then exchanges
// it for Graph API tokens. Refresh token handles ongoing renewal.
// ============================================================================

import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import http from "node:http";
import type { TokenData, IMsAuthProvider } from "./types.js";

const TOKEN_FILE = "~/.openclaw/unified-inbox-tokens.json";

// Microsoft Office Desktop public client ID (first-party, no app registration).
// No first-party client has Chat.Read preauthorized — Teams chat requires
// admin-granted consent via a custom Azure AD app registration.
const CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
const REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";
const TOKEN_ENDPOINT = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token";
const SCOPE = "https://graph.microsoft.com/.default offline_access openid profile";

// Token refresh 5 minutes before expiry
const REFRESH_BUFFER_MS = 5 * 60 * 1000;
// Auto-refresh check interval: 10 minutes
const AUTO_REFRESH_INTERVAL_MS = 10 * 60 * 1000;

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

export type MsAuthBrowserOptions = {
  browserProfile: string;
  cdpPort: number;
  tokenFile?: string;
};

/**
 * Browser-based auth provider. Opens Microsoft login in the Playwright browser,
 * captures the OAuth auth code via CDP, and exchanges it for Graph API tokens.
 * After initial auth, refresh tokens handle renewal without browser interaction.
 */
export class MsAuthBrowser implements IMsAuthProvider {
  private tokenData: TokenData | null = null;
  private refreshTimer: ReturnType<typeof setInterval> | null = null;
  private tokenFile: string;
  private browserProfile: string;
  private cdpPort: number;
  private log: Logger;

  constructor(opts: MsAuthBrowserOptions, log: Logger) {
    this.tokenFile = resolveHomePath(opts.tokenFile ?? TOKEN_FILE);
    this.browserProfile = opts.browserProfile;
    this.cdpPort = opts.cdpPort;
    this.log = log;
  }

  isAuthenticated(): boolean {
    return this.tokenData !== null;
  }

  async getAccessToken(): Promise<string> {
    if (!this.tokenData) {
      throw new Error("Not authenticated. Run /inbox_login first.");
    }

    // Refresh via refresh token if access token is near expiry
    if (Date.now() >= this.tokenData.expiresAt - REFRESH_BUFFER_MS) {
      const refreshed = await this.tryRefreshToken();
      if (!refreshed) {
        throw new Error(
          "Token expired and refresh failed. Run /inbox_login to re-authenticate.",
        );
      }
    }

    return this.tokenData.accessToken;
  }

  async loadPersistedTokens(): Promise<boolean> {
    try {
      const raw = await fs.readFile(this.tokenFile, "utf-8");
      const data = JSON.parse(raw) as TokenData;
      if (data.accessToken && data.expiresAt) {
        this.tokenData = data;
        // If access token expired, try refresh token
        if (Date.now() >= data.expiresAt) {
          return await this.tryRefreshToken();
        }
        return true;
      }
    } catch {
      // No token file or invalid — need fresh auth
    }
    return false;
  }

  startAutoRefresh(onError: (error: string) => Promise<void>): void {
    this.stopAutoRefresh();
    this.refreshTimer = setInterval(async () => {
      if (!this.tokenData) return;

      if (Date.now() >= this.tokenData.expiresAt - REFRESH_BUFFER_MS) {
        try {
          const refreshed = await this.tryRefreshToken();
          if (!refreshed) {
            await onError(
              "Token refresh failed. Use /inbox_login to re-authenticate.",
            );
          }
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          await onError(msg);
        }
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
  // Browser interaction methods (called from telegram-commands)
  // ---------------------------------------------------------------------------

  /**
   * Start the interactive auth flow: open the Microsoft login page in the
   * browser and wait for the user to complete sign-in. Returns the auth code
   * or null if the flow fails/times out.
   *
   * The flow:
   * 1. Navigate browser to OAuth authorize endpoint
   * 2. User sees "Are you trying to sign in to Microsoft Office?" → clicks Continue
   * 3. Browser redirects to urn:ietf:wg:oauth:2.0:oob?code=...
   * 4. We capture the code via CDP network interception
   */
  async startLoginFlow(timeoutMs = 5 * 60 * 1000): Promise<{ code: string } | { error: string }> {
    const authUrl = buildAuthUrl();

    try {
      const wsUrl = await this.getCdpWsUrl();
      return await this.captureAuthCode(wsUrl, authUrl, timeoutMs);
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  }

  /**
   * Exchange an auth code for access + refresh tokens.
   * Returns true if successful.
   */
  async exchangeAuthCode(code: string): Promise<boolean> {
    try {
      const tokenData = await exchangeCodeForTokens(code);
      this.tokenData = tokenData;
      await this.persistTokens();
      this.log.info("unified-inbox: Graph API tokens acquired and saved");
      return true;
    } catch (err) {
      this.log.error(
        `unified-inbox: token exchange failed: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  /**
   * Navigate the browser back to Outlook (after auth flow completes).
   */
  async navigateToOutlook(): Promise<void> {
    try {
      const wsUrl = await this.getCdpWsUrl();
      await this.cdpNavigate(wsUrl, "https://outlook.office.com/mail/");
    } catch {
      // Non-critical — the browser just shows the oob page
    }
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  private async tryRefreshToken(): Promise<boolean> {
    if (!this.tokenData?.refreshToken) return false;

    try {
      const body = new URLSearchParams({
        client_id: CLIENT_ID,
        grant_type: "refresh_token",
        refresh_token: this.tokenData.refreshToken,
        scope: SCOPE,
      });

      const resp = await fetch(TOKEN_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      });

      const data = await resp.json() as Record<string, unknown>;

      if (data.error) {
        this.log.error(`unified-inbox: refresh failed: ${data.error} - ${data.error_description}`);
        return false;
      }

      const accessToken = data.access_token as string;
      const refreshToken = (data.refresh_token as string) || this.tokenData.refreshToken;
      const expiresIn = data.expires_in as number;

      // Decode JWT for scopes and account
      const payload = decodeJwtPayload(accessToken);

      this.tokenData = {
        accessToken,
        refreshToken,
        expiresAt: Date.now() + expiresIn * 1000,
        scopes: (payload?.scp || "").split(" "),
        account: payload?.upn || payload?.unique_name || this.tokenData.account,
      };

      await this.persistTokens();
      this.log.info("unified-inbox: token refreshed successfully");
      return true;
    } catch (err) {
      this.log.error(
        `unified-inbox: token refresh error: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  /** Get the WebSocket debug URL for the browser's page target via CDP */
  private getCdpWsUrl(): Promise<string> {
    return new Promise((resolve, reject) => {
      http.get(`http://127.0.0.1:${this.cdpPort}/json`, (res) => {
        let data = "";
        res.on("data", (chunk: Buffer) => (data += chunk));
        res.on("end", () => {
          try {
            const targets = JSON.parse(data) as Array<{ type: string; url: string; webSocketDebuggerUrl: string }>;
            const page = targets.find((t) => t.type === "page");
            if (!page) reject(new Error("No browser tab found. Is the browser running?"));
            else resolve(page.webSocketDebuggerUrl);
          } catch {
            reject(new Error("Failed to parse CDP targets"));
          }
        });
      }).on("error", (err) => reject(new Error(`CDP connection failed (port ${this.cdpPort}): ${err.message}`)));
    });
  }

  /** Use CDP to navigate to auth URL and capture the oob redirect with auth code */
  private captureAuthCode(
    wsUrl: string,
    authUrl: string,
    timeoutMs: number,
  ): Promise<{ code: string } | { error: string }> {
    return new Promise((resolve) => {
      const ws = new WebSocket(wsUrl);
      let id = 1;

      const timer = setTimeout(() => {
        ws.close();
        resolve({ error: "Login timed out. Please try again." });
      }, timeoutMs);

      function send(method: string, params: Record<string, unknown> = {}) {
        ws.send(JSON.stringify({ id: id++, method, params }));
      }

      ws.addEventListener("open", () => {
        send("Network.enable");
        send("Page.enable");
        send("Page.navigate", { url: authUrl });
      });

      ws.addEventListener("message", (event: MessageEvent) => {
        const msg = JSON.parse(String(event.data)) as {
          method?: string;
          params?: { request?: { url?: string } };
        };

        // Capture redirect to native client URI containing the auth code
        if (msg.method === "Network.requestWillBeSent") {
          const url = msg.params?.request?.url || "";
          if (url.startsWith(REDIRECT_URI)) {
            clearTimeout(timer);
            send("Network.disable");
            ws.close();

            const codeMatch = url.match(/[?&]code=([^&]+)/);
            if (codeMatch) {
              resolve({ code: decodeURIComponent(codeMatch[1]) });
            } else {
              const errorMatch = url.match(/[?&]error_description=([^&]+)/);
              resolve({
                error: errorMatch
                  ? decodeURIComponent(errorMatch[1]).replace(/\+/g, " ")
                  : "Authentication failed — no code received",
              });
            }
          }
        }
      });

      ws.addEventListener("error", () => {
        clearTimeout(timer);
        resolve({ error: "WebSocket connection to browser failed" });
      });
    });
  }

  /** Navigate the browser to a URL via CDP */
  private cdpNavigate(wsUrl: string, url: string): Promise<void> {
    return new Promise((resolve) => {
      const ws = new WebSocket(wsUrl);
      ws.addEventListener("open", () => {
        ws.send(JSON.stringify({ id: 1, method: "Page.navigate", params: { url } }));
        setTimeout(() => { ws.close(); resolve(); }, 2000);
      });
      ws.addEventListener("error", () => resolve());
    });
  }

  private async persistTokens(): Promise<void> {
    if (!this.tokenData) return;

    const dir = path.dirname(this.tokenFile);
    await fs.mkdir(dir, { recursive: true });
    await fs.writeFile(
      this.tokenFile,
      JSON.stringify(this.tokenData, null, 2),
      { mode: 0o600 },
    );
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function buildAuthUrl(): string {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    scope: SCOPE,
    response_mode: "query",
  });
  return `https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?${params}`;
}

async function exchangeCodeForTokens(code: string): Promise<TokenData> {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    grant_type: "authorization_code",
    code,
    redirect_uri: REDIRECT_URI,
    scope: SCOPE,
  });

  const resp = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const data = await resp.json() as Record<string, unknown>;

  if (data.error) {
    throw new Error(`${data.error}: ${data.error_description}`);
  }

  const accessToken = data.access_token as string;
  const payload = decodeJwtPayload(accessToken);

  return {
    accessToken,
    refreshToken: (data.refresh_token as string) || "",
    expiresAt: Date.now() + (data.expires_in as number) * 1000,
    scopes: (payload?.scp || "").split(" "),
    account: payload?.upn || payload?.unique_name || undefined,
  };
}

function decodeJwtPayload(token: string): Record<string, string> | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    return JSON.parse(Buffer.from(parts[1], "base64url").toString());
  } catch {
    return null;
  }
}

function resolveHomePath(p: string): string {
  if (p.startsWith("~/")) {
    return path.join(os.homedir(), p.slice(2));
  }
  return p;
}
