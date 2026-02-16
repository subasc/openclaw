// ============================================================================
// Browser-based Microsoft auth: extracts MSAL tokens from Outlook Web via
// OpenClaw's browser control HTTP API (Playwright persistent profiles)
// ============================================================================

import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import type { TokenData, IMsAuthProvider } from "./types.js";

const TOKEN_FILE = "~/.openclaw/unified-inbox-tokens.json";
const OUTLOOK_URL = "https://outlook.office.com";
const GRAPH_SCOPE_PREFIX = "https://graph.microsoft.com/";

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
  browserApiPort: number;
  browserProfile: string;
  tokenFile?: string;
};

/**
 * Browser-based auth provider. Extracts MSAL tokens that Outlook Web
 * stores in localStorage after the user signs in with MFA.
 */
export class MsAuthBrowser implements IMsAuthProvider {
  private tokenData: TokenData | null = null;
  private refreshTimer: ReturnType<typeof setInterval> | null = null;
  private tokenFile: string;
  private browserApiBase: string;
  private browserProfile: string;
  private log: Logger;

  constructor(opts: MsAuthBrowserOptions, log: Logger) {
    this.tokenFile = resolveHomePath(opts.tokenFile ?? TOKEN_FILE);
    this.browserApiBase = `http://127.0.0.1:${opts.browserApiPort}`;
    this.browserProfile = opts.browserProfile;
    this.log = log;
  }

  isAuthenticated(): boolean {
    return this.tokenData !== null;
  }

  async getAccessToken(): Promise<string> {
    if (!this.tokenData) {
      throw new Error("Not authenticated. Run /inbox_login first.");
    }

    // Re-extract from browser if token is near expiry
    if (Date.now() >= this.tokenData.expiresAt - REFRESH_BUFFER_MS) {
      const refreshed = await this.tryRefreshFromBrowser();
      if (!refreshed) {
        throw new Error(
          "Token expired and browser re-extraction failed. Run /inbox_login to re-authenticate.",
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
        // If expired, try refreshing from browser
        if (Date.now() >= data.expiresAt) {
          return await this.tryRefreshFromBrowser();
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

      // Check if token needs refresh
      if (Date.now() >= this.tokenData.expiresAt - REFRESH_BUFFER_MS) {
        try {
          const refreshed = await this.tryRefreshFromBrowser();
          if (!refreshed) {
            await onError(
              "Browser token re-extraction failed. Session may have expired. Use /inbox_login to re-authenticate.",
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

  /** Open the Outlook Web login page in the browser */
  async openLoginPage(): Promise<void> {
    await this.browserNavigate(OUTLOOK_URL);
  }

  /**
   * Extract MSAL tokens from browser localStorage.
   * Returns true if valid Graph API tokens were found.
   */
  async extractTokensFromBrowser(): Promise<boolean> {
    try {
      const storage = await this.browserGetLocalStorage();
      if (!storage) return false;

      const tokenData = parseMsalTokensFromStorage(storage);
      if (!tokenData) return false;

      this.tokenData = tokenData;
      await this.persistTokens();
      this.log.info("unified-inbox: browser tokens extracted and saved");
      return true;
    } catch (err) {
      this.log.error(
        `unified-inbox: browser token extraction failed: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  // ---------------------------------------------------------------------------
  // Private: browser API helpers
  // ---------------------------------------------------------------------------

  private async tryRefreshFromBrowser(): Promise<boolean> {
    try {
      return await this.extractTokensFromBrowser();
    } catch {
      return false;
    }
  }

  /** Navigate the browser to a URL via the browser control API */
  private async browserNavigate(url: string): Promise<void> {
    const resp = await fetch(`${this.browserApiBase}/navigate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url, profile: this.browserProfile }),
    });

    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      throw new Error(
        `Browser navigate failed (${resp.status}): ${text}`,
      );
    }
  }

  /** Get all localStorage entries from the current page via browser control API */
  private async browserGetLocalStorage(): Promise<Record<string, string> | null> {
    const resp = await fetch(
      `${this.browserApiBase}/storage/local?profile=${encodeURIComponent(this.browserProfile)}&origin=${encodeURIComponent(OUTLOOK_URL)}`,
    );

    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      this.log.warn(
        `unified-inbox: browser storage read failed (${resp.status}): ${text}`,
      );
      return null;
    }

    return (await resp.json()) as Record<string, string>;
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
// MSAL token parsing from localStorage
// ---------------------------------------------------------------------------

/**
 * Parse MSAL.js cache entries from browser localStorage to find
 * a valid access token with Microsoft Graph scopes.
 *
 * MSAL.js stores tokens in localStorage with keys like:
 * - Access tokens: keys containing "accesstoken" (JSON with secret, expiresOn, target)
 * - Refresh tokens: keys containing "refreshtoken" (JSON with secret)
 * - Account info: keys containing "account" (JSON with username, etc.)
 */
function parseMsalTokensFromStorage(
  storage: Record<string, string>,
): TokenData | null {
  let bestToken: {
    accessToken: string;
    expiresAt: number;
    scopes: string[];
  } | null = null;

  let refreshToken = "";
  let account = "";

  for (const [key, value] of Object.entries(storage)) {
    const keyLower = key.toLowerCase();

    // Look for access tokens with Graph scopes
    if (keyLower.includes("accesstoken")) {
      try {
        const entry = JSON.parse(value) as {
          secret?: string;
          expires_on?: string;
          extended_expires_on?: string;
          target?: string;
        };

        if (!entry.secret || !entry.target) continue;

        // Check if this token has Graph API scopes
        const scopes = entry.target.split(" ");
        const hasGraphScope = scopes.some(
          (s) =>
            s.startsWith(GRAPH_SCOPE_PREFIX) ||
            s.toLowerCase().startsWith("mail.") ||
            s.toLowerCase().startsWith("calendars.") ||
            s.toLowerCase().startsWith("chat.") ||
            s.toLowerCase().startsWith("user."),
        );

        if (!hasGraphScope) continue;

        const expiresOn = parseInt(entry.expires_on ?? "0", 10) * 1000;

        // Skip expired tokens
        if (expiresOn <= Date.now()) continue;

        // Prefer token with latest expiry
        if (!bestToken || expiresOn > bestToken.expiresAt) {
          bestToken = {
            accessToken: entry.secret,
            expiresAt: expiresOn,
            scopes,
          };
        }
      } catch {
        // Skip malformed entries
      }
    }

    // Look for refresh tokens
    if (keyLower.includes("refreshtoken") && !refreshToken) {
      try {
        const entry = JSON.parse(value) as { secret?: string };
        if (entry.secret) {
          refreshToken = entry.secret;
        }
      } catch {
        // Skip
      }
    }

    // Look for account info
    if (
      keyLower.includes("account") &&
      !keyLower.includes("token") &&
      !account
    ) {
      try {
        const entry = JSON.parse(value) as { username?: string };
        if (entry.username) {
          account = entry.username;
        }
      } catch {
        // Skip
      }
    }
  }

  if (!bestToken) return null;

  return {
    accessToken: bestToken.accessToken,
    refreshToken,
    expiresAt: bestToken.expiresAt,
    account: account || undefined,
    scopes: bestToken.scopes,
  };
}

function resolveHomePath(p: string): string {
  if (p.startsWith("~/")) {
    return path.join(os.homedir(), p.slice(2));
  }
  return p;
}
