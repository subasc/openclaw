// ============================================================================
// Microsoft OAuth2 Device Code Flow authentication
// Uses Microsoft Office Desktop public client ID — no app registration needed
// ============================================================================

import {
  PublicClientApplication,
  type AuthenticationResult,
  type DeviceCodeRequest,
} from "@azure/msal-node";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import type { TokenData } from "./types.js";

const SCOPES = [
  "Mail.Read",
  "Mail.Send",
  "Calendars.Read",
  "Chat.Read",
  "Chat.ReadWrite",
  "User.Read",
  "offline_access",
];

export type DeviceCodeCallback = (message: string) => void | Promise<void>;

export type MsAuthOptions = {
  clientId: string;
  tenantId: string;
  tokenFile: string;
};

/**
 * Microsoft auth provider using device code flow.
 * Manages token acquisition, persistence, and refresh.
 */
export class MsAuth {
  private pca: PublicClientApplication;
  private tokenFile: string;
  private tokenData: TokenData | null = null;
  private refreshTimer: ReturnType<typeof setInterval> | null = null;

  constructor(private opts: MsAuthOptions) {
    this.pca = new PublicClientApplication({
      auth: {
        clientId: opts.clientId,
        authority: `https://login.microsoftonline.com/${opts.tenantId}`,
      },
    });
    this.tokenFile = resolveHomePath(opts.tokenFile);
  }

  /** Load persisted tokens from disk. Returns true if valid tokens found. */
  async loadPersistedTokens(): Promise<boolean> {
    try {
      const raw = await fs.readFile(this.tokenFile, "utf-8");
      const data = JSON.parse(raw) as TokenData;
      if (data.accessToken && data.refreshToken && data.expiresAt) {
        this.tokenData = data;
        // If expired, try refreshing
        if (Date.now() >= data.expiresAt) {
          return await this.refreshAccessToken();
        }
        return true;
      }
    } catch {
      // No token file or invalid — need fresh auth
    }
    return false;
  }

  /**
   * Start device code flow. The callback receives the "go to URL and enter code" message
   * which should be sent to the user via Telegram.
   */
  async authenticateWithDeviceCode(
    onDeviceCode: DeviceCodeCallback,
  ): Promise<boolean> {
    const request: DeviceCodeRequest = {
      scopes: SCOPES,
      deviceCodeCallback: async (response) => {
        await onDeviceCode(response.message);
      },
    };

    try {
      const result = await this.pca.acquireTokenByDeviceCode(request);
      if (result) {
        await this.storeAuthResult(result);
        return true;
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(`Device code authentication failed: ${msg}`);
    }
    return false;
  }

  /** Get a valid access token, refreshing if needed. */
  async getAccessToken(): Promise<string> {
    if (!this.tokenData) {
      throw new Error("Not authenticated. Run /inbox-login first.");
    }

    // Refresh if token expires within 5 minutes
    if (Date.now() >= this.tokenData.expiresAt - 5 * 60 * 1000) {
      const refreshed = await this.refreshAccessToken();
      if (!refreshed) {
        throw new Error(
          "Token refresh failed. Run /inbox-login to re-authenticate.",
        );
      }
    }

    return this.tokenData.accessToken;
  }

  /** Check if we have valid (or refreshable) tokens */
  isAuthenticated(): boolean {
    return this.tokenData !== null;
  }

  /** Start the auto-refresh loop (every 45 minutes) */
  startAutoRefresh(onError?: (error: string) => void): void {
    this.stopAutoRefresh();
    this.refreshTimer = setInterval(
      async () => {
        try {
          await this.refreshAccessToken();
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          onError?.(msg);
        }
      },
      45 * 60 * 1000,
    );
  }

  /** Stop the auto-refresh loop */
  stopAutoRefresh(): void {
    if (this.refreshTimer) {
      clearInterval(this.refreshTimer);
      this.refreshTimer = null;
    }
  }

  // ---------------------------------------------------------------------------
  // Private
  // ---------------------------------------------------------------------------

  private async refreshAccessToken(): Promise<boolean> {
    if (!this.tokenData?.refreshToken) return false;

    try {
      // MSAL node's silent acquisition uses the token cache.
      // For device code flow with public client, we use refresh_token grant directly.
      const result = await this.pca.acquireTokenByRefreshToken({
        refreshToken: this.tokenData.refreshToken,
        scopes: SCOPES,
      });

      if (result) {
        await this.storeAuthResult(result);
        return true;
      }
    } catch {
      // Refresh token expired or revoked
      this.tokenData = null;
    }
    return false;
  }

  private async storeAuthResult(result: AuthenticationResult): Promise<void> {
    this.tokenData = {
      accessToken: result.accessToken,
      refreshToken: (result as Record<string, unknown>).refreshToken as string ?? this.tokenData?.refreshToken ?? "",
      expiresAt: result.expiresOn?.getTime() ?? Date.now() + 60 * 60 * 1000,
      account: result.account?.username,
      scopes: result.scopes,
    };

    // Persist to disk
    const dir = path.dirname(this.tokenFile);
    await fs.mkdir(dir, { recursive: true });
    await fs.writeFile(this.tokenFile, JSON.stringify(this.tokenData, null, 2), {
      mode: 0o600, // Owner read/write only
    });
  }
}

function resolveHomePath(p: string): string {
  if (p.startsWith("~/")) {
    return path.join(os.homedir(), p.slice(2));
  }
  return p;
}
