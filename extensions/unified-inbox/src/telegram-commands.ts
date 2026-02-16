// ============================================================================
// Telegram slash commands for the Unified Inbox
// ============================================================================

import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import type { UnifiedInboxConfig } from "./config.js";
import type { MsAuth } from "./ms-auth.js";
import type { EmailMonitor } from "./email-monitor.js";
import type { CalendarMonitor } from "./calendar-monitor.js";
import type { TeamsChatMonitor } from "./teams-chat-monitor.js";
import { fetchMailDelta, sendMail } from "./ms-graph-client.js";
import {
  formatEmailListItem,
  formatCalendarListItem,
} from "./formatters.js";

// Shared state — set by service.ts after initialization
let sharedAuth: MsAuth | null = null;
let sharedEmailMonitor: EmailMonitor | null = null;
let sharedCalendarMonitor: CalendarMonitor | null = null;
let sharedTeamsChatMonitor: TeamsChatMonitor | null = null;

export function setCommandDependencies(deps: {
  auth: MsAuth;
  emailMonitor?: EmailMonitor;
  calendarMonitor?: CalendarMonitor;
  teamsChatMonitor?: TeamsChatMonitor;
}): void {
  sharedAuth = deps.auth;
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
      if (!sharedAuth?.isAuthenticated()) {
        return { text: "Not authenticated. Use /inbox_login first." };
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

  // /inbox_login — re-authenticate with Microsoft
  api.registerCommand({
    name: "inbox_login",
    description: "Authenticate with Microsoft 365 (device code flow)",
    handler: async (_ctx) => {
      if (!sharedAuth) {
        return { text: "Unified inbox service not initialized." };
      }

      try {
        // This will trigger the device code callback which sends the code to Telegram
        const success = await sharedAuth.authenticateWithDeviceCode(
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
    },
  });

  // /inbox_status — show monitor statuses
  api.registerCommand({
    name: "inbox_status",
    description: "Show Unified Inbox monitor statuses",
    handler: async (_ctx) => {
      const authenticated = sharedAuth?.isAuthenticated() ?? false;

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
        `Authenticated: ${authenticated ? "yes" : "NO"}`,
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

function esc(text: string): string {
  return text.replace(/([_*\[\]()~`>#+\-=|{}.!\\])/g, "\\$1");
}
