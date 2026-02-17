// ============================================================================
// Telegram slash commands for the Unified Inbox
// ============================================================================

import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import type { UnifiedInboxConfig } from "./config.js";
import type { IMsAuthProvider } from "./types.js";
import type { MsAuth } from "./ms-auth.js";
import type { MsAuthBrowser } from "./ms-auth-browser.js";
import type { EmailMonitor } from "./email-monitor.js";
import type { CalendarMonitor } from "./calendar-monitor.js";
import type { TeamsChatMonitor } from "./teams-chat-monitor.js";
import { fetchMailDelta, sendMail } from "./ms-graph-client.js";
import {
  formatEmailListItem,
  formatCalendarListItem,
} from "./formatters.js";

// Shared state — set by service.ts after initialization
let sharedAuth: IMsAuthProvider | null = null;
let sharedTeamsAuth: IMsAuthProvider | null = null;
let sharedOutlookRestAuth: IMsAuthProvider | null = null;
let sharedAuthMode: "browser" | "device-code" = "browser";
let sharedEmailMonitor: EmailMonitor | null = null;
let sharedCalendarMonitor: CalendarMonitor | null = null;
let sharedTeamsChatMonitor: TeamsChatMonitor | null = null;

export function setCommandDependencies(deps: {
  auth: IMsAuthProvider;
  teamsAuth?: IMsAuthProvider | null;
  outlookRestAuth?: IMsAuthProvider | null;
  authMode?: "browser" | "device-code";
  emailMonitor?: EmailMonitor;
  calendarMonitor?: CalendarMonitor;
  teamsChatMonitor?: TeamsChatMonitor;
}): void {
  sharedAuth = deps.auth;
  sharedTeamsAuth = deps.teamsAuth ?? null;
  sharedOutlookRestAuth = deps.outlookRestAuth ?? null;
  sharedAuthMode = deps.authMode ?? "browser";
  sharedEmailMonitor = deps.emailMonitor ?? null;
  sharedCalendarMonitor = deps.calendarMonitor ?? null;
  sharedTeamsChatMonitor = deps.teamsChatMonitor ?? null;
}

export function registerInboxCommands(
  api: OpenClawPluginApi,
  cfg: UnifiedInboxConfig,
): void {
  // /inbox — show recent unread emails
  api.registerCommand({
    name: "inbox",
    description: "Show recent unread emails",
    handler: async (_ctx) => {
      if (!sharedAuth?.isAuthenticated()) {
        return { text: "Not authenticated. Use /inbox_login first." };
      }

      try {
        const token = await sharedAuth.getAccessToken();
        const result = await fetchMailDelta(token, "Inbox", {
          filterUnread: true,
          top: 5,
        });

        if (result.messages.length === 0) {
          return { text: "No unread emails." };
        }

        const lines = result.messages
          .slice(0, 5)
          .map((msg, i) => formatEmailListItem(msg, i));

        return {
          text: `*Unread Emails:*\n\n${lines.join("\n")}`,
          parseMode: "MarkdownV2",
        };
      } catch (err) {
        return { text: `Error: ${err instanceof Error ? err.message : String(err)}` };
      }
    },
  });

  // /calendar — show today's upcoming events
  api.registerCommand({
    name: "calendar",
    description: "Show today's upcoming events",
    handler: async (_ctx) => {
      if (!sharedAuth?.isAuthenticated()) {
        return { text: "Not authenticated. Use /inbox_login first." };
      }

      try {
        if (!sharedCalendarMonitor) {
          return { text: "Calendar monitor not running." };
        }

        const events = await sharedCalendarMonitor.fetchUpcomingEvents();

        if (events.length === 0) {
          return { text: "No more events today." };
        }

        const lines = events
          .filter((e) => !e.isCancelled)
          .map((event, i) => formatCalendarListItem(event, i));

        return {
          text: `*Today's Events:*\n\n${lines.join("\n")}`,
          parseMode: "MarkdownV2",
        };
      } catch (err) {
        return { text: `Error: ${err instanceof Error ? err.message : String(err)}` };
      }
    },
  });

  // /teams — show recent Teams chat activity
  api.registerCommand({
    name: "teams",
    description: "Show recent Teams chat activity",
    handler: async (_ctx) => {
      if (!sharedTeamsAuth?.isAuthenticated()) {
        return { text: "Teams not authenticated. Is the browser running and logged into Teams?" };
      }

      try {
        if (!sharedTeamsChatMonitor) {
          return { text: "Teams chat monitor not running." };
        }

        const chats = await sharedTeamsChatMonitor.fetchRecentChats();

        if (chats.length === 0) {
          return { text: "No recent Teams chats." };
        }

        const lines = chats
          .filter((c) => c.lastMessagePreview)
          .slice(0, 10)
          .map((chat, i) => {
            const name = chat.topic || chat.chatType;
            const sender =
              chat.lastMessagePreview?.from?.user?.displayName ||
              chat.lastMessagePreview?.from?.application?.displayName ||
              "Unknown";
            const preview = (chat.lastMessagePreview?.body?.content || "")
              .replace(/<[^>]+>/g, "")
              .slice(0, 50);
            return `${i + 1}. *${esc(name)}* - ${esc(sender)}: ${esc(preview)}`;
          });

        return {
          text: `*Recent Teams Chats:*\n\n${lines.join("\n")}`,
          parseMode: "MarkdownV2",
        };
      } catch (err) {
        return { text: `Error: ${err instanceof Error ? err.message : String(err)}` };
      }
    },
  });

  // /compose <to> <subject> — compose a new email
  api.registerCommand({
    name: "compose",
    description: "Compose a new email: /compose to@email.com Subject here",
    acceptsArgs: true,
    handler: async (ctx) => {
      if (!sharedAuth?.isAuthenticated()) {
        return { text: "Not authenticated. Use /inbox_login first." };
      }

      const args = ctx.args?.trim();
      if (!args) {
        return {
          text: "Usage: /compose recipient@email.com Subject of the email\nThen type your message body.",
        };
      }

      // Parse: first word is email, rest is subject
      const parts = args.split(/\s+/);
      const to = parts[0];
      const subject = parts.slice(1).join(" ");

      if (!to?.includes("@") || !subject) {
        return {
          text: "Usage: /compose recipient@email.com Subject of the email",
        };
      }

      // For now, send a placeholder. A more interactive flow could be added later.
      try {
        const token = await sharedAuth.getAccessToken();
        await sendMail(token, {
          to,
          subject,
          body: "(Composed via OpenClaw Unified Inbox)",
        });
        return { text: `Email sent to ${to} with subject: ${subject}` };
      } catch (err) {
        return { text: `Failed to send: ${err instanceof Error ? err.message : String(err)}` };
      }
    },
  });

  // /refresh_tokens — force re-extract/recover tokens from browser tabs
  api.registerCommand({
    name: "refresh_tokens",
    description: "Force refresh Microsoft 365 tokens",
    handler: async (_ctx) => {
      const providers: Array<{ name: string; auth: IMsAuthProvider | null }> = [
        { name: "Outlook", auth: sharedAuth },
        { name: "Teams", auth: sharedTeamsAuth },
        { name: "Outlook-REST", auth: sharedOutlookRestAuth },
      ];

      const results: string[] = [];
      for (const { name, auth } of providers) {
        if (!auth) {
          results.push(`${name}: not configured`);
          continue;
        }
        if (typeof auth.forceRefresh !== "function") {
          results.push(`${name}: forceRefresh not supported`);
          continue;
        }
        const ok = await auth.forceRefresh();
        results.push(`${name}: ${ok ? "refreshed" : "FAILED"}`);
      }

      // Auto-unpause any paused monitors
      const monitors = [
        { name: "Email", monitor: sharedEmailMonitor },
        { name: "Calendar", monitor: sharedCalendarMonitor },
        { name: "Teams Chat", monitor: sharedTeamsChatMonitor },
      ];
      const unpaused: string[] = [];
      for (const { name, monitor } of monitors) {
        if (monitor?.status?.paused) {
          monitor.status.paused = false;
          monitor.status.consecutiveFailures = 0;
          unpaused.push(name);
        }
      }

      let text = `Token refresh results:\n${results.join("\n")}`;
      if (unpaused.length > 0) {
        text += `\n\nUnpaused monitors: ${unpaused.join(", ")}`;
      }

      return { text };
    },
  });

  // /restart_gateway — restart the OpenClaw gateway remotely
  api.registerCommand({
    name: "restart_gateway",
    description: "Restart the OpenClaw gateway",
    handler: async (_ctx) => {
      // Schedule SIGTERM restart with a short delay so the Telegram response is sent first.
      // launchd/systemd will auto-restart the process after SIGTERM.
      setTimeout(() => {
        process.kill(process.pid, "SIGTERM");
      }, 1_500);

      return {
        text: "Gateway restart scheduled. The process will terminate in ~2 seconds and auto-restart via launchd.",
      };
    },
  });

  // /inbox_login — authenticate with Microsoft
  api.registerCommand({
    name: "inbox_login",
    description: "Authenticate with Microsoft 365",
    handler: async (_ctx) => {
      if (!sharedAuth) {
        return { text: "Unified inbox service not initialized." };
      }

      if (sharedAuthMode === "browser") {
        return await handleBrowserLogin(cfg);
      }

      return await handleDeviceCodeLogin(cfg);
    },
  });

  // /inbox_status — show monitor statuses
  api.registerCommand({
    name: "inbox_status",
    description: "Show Unified Inbox monitor statuses",
    handler: async (_ctx) => {
      const outlookAuth = sharedAuth?.isAuthenticated() ?? false;
      const teamsAuth = sharedTeamsAuth?.isAuthenticated() ?? false;

      const formatStatus = (name: string, status: { running: boolean; paused: boolean; lastPollAt: number | null; lastError: string | null } | null): string => {
        if (!status) return `${name}: not configured`;
        const state = status.paused
          ? "PAUSED"
          : status.running
            ? "running"
            : "stopped";
        const lastPoll = status.lastPollAt
          ? new Date(status.lastPollAt).toLocaleTimeString()
          : "never";
        const error = status.lastError
          ? `\n  Last error: ${status.lastError.slice(0, 80)}`
          : "";
        return `${name}: ${state} (last poll: ${lastPoll})${error}`;
      };

      const lines = [
        `Auth mode: CDP token extraction`,
        `Outlook auth: ${outlookAuth ? "yes" : "NO"}`,
        `Teams auth: ${teamsAuth ? "yes" : "NO"}`,
        "",
        formatStatus("Email", sharedEmailMonitor?.status ?? null),
        formatStatus("Calendar", sharedCalendarMonitor?.status ?? null),
        formatStatus("Teams Chat", sharedTeamsChatMonitor?.status ?? null),
        `WhatsApp Bridge: ${cfg.whatsapp.enabled ? "enabled" : "disabled"}`,
      ];

      return { text: lines.join("\n") };
    },
  });
}

// ---------------------------------------------------------------------------
// Browser-based login flow
// ---------------------------------------------------------------------------

async function handleBrowserLogin(
  cfg: UnifiedInboxConfig,
): Promise<{ text: string }> {
  const browserAuth = sharedAuth as MsAuthBrowser;

  const { sendTelegramMessage } = await import("./telegram-sender.js");
  await sendTelegramMessage({
    botToken: cfg.telegramBotToken,
    chatId: cfg.telegramChatId,
    text: 'Opening Microsoft login in browser... Click "Continue" when prompted.',
  });

  // Start the OAuth auth code flow via CDP
  const result = await browserAuth.startLoginFlow(5 * 60 * 1000);

  if ("error" in result) {
    return {
      text: `Login failed: ${result.error}\nMake sure the browser profile is running.`,
    };
  }

  // Exchange auth code for tokens
  await sendTelegramMessage({
    botToken: cfg.telegramBotToken,
    chatId: cfg.telegramChatId,
    text: "Sign-in confirmed. Exchanging for Graph API tokens...",
  });

  const success = await browserAuth.exchangeAuthCode(result.code);
  if (!success) {
    return { text: "Token exchange failed. Try /inbox_login again." };
  }

  // Navigate browser back to Outlook
  await browserAuth.navigateToOutlook();

  return {
    text: "Authentication successful! Graph API tokens acquired. Monitors will start automatically.",
  };
}

// ---------------------------------------------------------------------------
// Device-code login flow (original)
// ---------------------------------------------------------------------------

async function handleDeviceCodeLogin(
  cfg: UnifiedInboxConfig,
): Promise<{ text: string }> {
  const deviceCodeAuth = sharedAuth as MsAuth;

  try {
    const success = await deviceCodeAuth.authenticateWithDeviceCode(
      async (message) => {
        const { sendTelegramMessage } = await import("./telegram-sender.js");
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: message,
        });
      },
    );

    return {
      text: success
        ? "Authentication successful! Monitors will start automatically."
        : "Authentication failed. Try again with /inbox_login.",
    };
  } catch (err) {
    return { text: `Auth error: ${err instanceof Error ? err.message : String(err)}` };
  }
}

function esc(text: string): string {
  return text.replace(/([_*\[\]()~`>#+\-=|{}.!\\])/g, "\\$1");
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
