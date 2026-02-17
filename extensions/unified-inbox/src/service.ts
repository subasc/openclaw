// ============================================================================
// Unified Inbox service lifecycle
// Manages all monitors, token refresh, and graceful shutdown
// ============================================================================

import type {
  OpenClawPluginService,
  OpenClawPluginServiceContext,
} from "openclaw/plugin-sdk";
import type { UnifiedInboxConfig } from "./config.js";
import type { IMsAuthProvider } from "./types.js";
import { BrowserTokenAuth } from "./browser-token-auth.js";
import { ReplyStore } from "./reply-store.js";
import { EmailMonitor } from "./email-monitor.js";
import { CalendarMonitor } from "./calendar-monitor.js";
import { TeamsChatMonitor } from "./teams-chat-monitor.js";
import { setWhatsAppBridgeReplyStore } from "./whatsapp-bridge.js";
import { setReplyRouterDependencies } from "./reply-router.js";
import { setCommandDependencies } from "./telegram-commands.js";
import { sendTelegramMessage } from "./telegram-sender.js";

type Logger = { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void };

export function createUnifiedInboxService(
  cfg: UnifiedInboxConfig,
  log: Logger,
): OpenClawPluginService {
  let auth: IMsAuthProvider | null = null;
  let teamsAuth: IMsAuthProvider | null = null;
  let replyStore: ReplyStore | null = null;
  let emailMonitor: EmailMonitor | null = null;
  let calendarMonitor: CalendarMonitor | null = null;
  let teamsChatMonitor: TeamsChatMonitor | null = null;

  return {
    id: "unified-inbox",

    async start(_ctx: OpenClawPluginServiceContext): Promise<void> {
      log.info("unified-inbox: service starting...");

      // 1. Initialize auth providers — CDP token extraction from live browser tabs
      const outlookAuth = new BrowserTokenAuth(
        { cdpPort: cfg.browserCdpPort, pageUrl: "https://outlook.office.com", pageName: "Outlook" },
        log,
      );
      auth = outlookAuth;

      if (cfg.teamsChat.enabled) {
        teamsAuth = new BrowserTokenAuth(
          { cdpPort: cfg.browserCdpPort, pageUrl: "https://teams.microsoft.com", pageName: "Teams" },
          log,
        );
      }

      log.info("unified-inbox: using CDP token extraction (Outlook" + (teamsAuth ? " + Teams" : "") + ")");

      // 2. Initialize reply store
      replyStore = new ReplyStore({
        storeFile: cfg.replyTracking.storeFile,
        maxEntries: cfg.replyTracking.maxEntries,
        ttlMs: cfg.replyTracking.ttlMs,
      });
      await replyStore.load();
      replyStore.startAutoFlush();

      // 3. Wire up shared dependencies
      setWhatsAppBridgeReplyStore(replyStore);
      setReplyRouterDependencies({ auth, teamsAuth: teamsAuth ?? auth, replyStore });

      // 4. Try loading tokens from browser tabs
      const hasOutlookTokens = await auth.loadPersistedTokens();
      const hasTeamsTokens = teamsAuth
        ? await teamsAuth.loadPersistedTokens().catch(() => false)
        : false;

      if (!hasOutlookTokens && !hasTeamsTokens) {
        log.info(
          "unified-inbox: no tokens found. Is the browser running and logged into Microsoft 365?",
        );
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: "[Unified Inbox] Started but no tokens found. Make sure the browser profile is running and you're logged into Microsoft 365.",
        });

        // Still wire up commands so /inbox_status works
        setCommandDependencies({ auth, teamsAuth, authMode: cfg.authMode });
        return;
      }

      // 5. Start auto-refresh for token re-extraction
      if (hasOutlookTokens) {
        auth.startAutoRefresh(async (error) => {
          log.error(`unified-inbox: Outlook token refresh failed: ${error}`);
          await sendTelegramMessage({
            botToken: cfg.telegramBotToken,
            chatId: cfg.telegramChatId,
            text: `[Unified Inbox] Outlook token refresh failed: ${error}`,
          });
        });
      }

      if (hasTeamsTokens && teamsAuth) {
        teamsAuth.startAutoRefresh(async (error) => {
          log.error(`unified-inbox: Teams token refresh failed: ${error}`);
          await sendTelegramMessage({
            botToken: cfg.telegramBotToken,
            chatId: cfg.telegramChatId,
            text: `[Unified Inbox] Teams token refresh failed: ${error}`,
          });
        });
      }

      // 6. Start monitors with their respective auth providers
      if (cfg.email.enabled && hasOutlookTokens) {
        emailMonitor = new EmailMonitor(cfg, auth, replyStore, log);
        await emailMonitor.start();
      }

      if (cfg.calendar.enabled && hasOutlookTokens) {
        calendarMonitor = new CalendarMonitor(cfg, auth, log);
        await calendarMonitor.start();
      }

      if (cfg.teamsChat.enabled && hasTeamsTokens && teamsAuth) {
        teamsChatMonitor = new TeamsChatMonitor(cfg, teamsAuth, replyStore, log);
        await teamsChatMonitor.start();
      }

      // 7. Wire up command dependencies (with monitors)
      setCommandDependencies({
        auth,
        teamsAuth,
        authMode: cfg.authMode,
        emailMonitor: emailMonitor ?? undefined,
        calendarMonitor: calendarMonitor ?? undefined,
        teamsChatMonitor: teamsChatMonitor ?? undefined,
      });

      // 8. Notify startup
      const monitors = [
        emailMonitor ? "Email" : null,
        calendarMonitor ? "Calendar" : null,
        teamsChatMonitor ? "Teams Chat" : null,
        cfg.whatsapp.enabled ? "WhatsApp Bridge" : null,
      ]
        .filter(Boolean)
        .join(", ");

      log.info(`unified-inbox: monitors started (${monitors || "none"})`);
      await sendTelegramMessage({
        botToken: cfg.telegramBotToken,
        chatId: cfg.telegramChatId,
        text: `[Unified Inbox] Started (CDP token extraction). Active monitors: ${monitors || "none"}`,
      });
    },

    async stop(_ctx: OpenClawPluginServiceContext): Promise<void> {
      log.info("unified-inbox: service stopping...");

      emailMonitor?.stop();
      calendarMonitor?.stop();
      teamsChatMonitor?.stop();
      auth?.stopAutoRefresh();
      teamsAuth?.stopAutoRefresh();

      if (replyStore) {
        replyStore.stopAutoFlush();
        await replyStore.flush();
      }

      log.info("unified-inbox: service stopped");
    },
  };
}
