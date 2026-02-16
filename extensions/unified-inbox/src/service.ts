// ============================================================================
// Unified Inbox service lifecycle
// Manages all monitors, token refresh, and graceful shutdown
// ============================================================================

import type {
  OpenClawPluginService,
  OpenClawPluginServiceContext,
} from "openclaw/plugin-sdk";
import type { UnifiedInboxConfig } from "./config.js";
import { MsAuth } from "./ms-auth.js";
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
  let auth: MsAuth | null = null;
  let replyStore: ReplyStore | null = null;
  let emailMonitor: EmailMonitor | null = null;
  let calendarMonitor: CalendarMonitor | null = null;
  let teamsChatMonitor: TeamsChatMonitor | null = null;

  return {
    id: "unified-inbox",

    async start(_ctx: OpenClawPluginServiceContext): Promise<void> {
      log.info("unified-inbox: service starting...");

      // 1. Initialize token store
      auth = new MsAuth({
        clientId: cfg.microsoft.clientId,
        tenantId: cfg.microsoft.tenantId,
        tokenFile: cfg.microsoft.tokenFile,
      });

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
      setReplyRouterDependencies({ auth, replyStore });

      // 4. Try loading persisted tokens
      const hasTokens = await auth.loadPersistedTokens();

      if (!hasTokens) {
        log.info(
          "unified-inbox: no valid tokens found. Use /inbox_login to authenticate.",
        );
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: "[Unified Inbox] Started but not authenticated. Send /inbox_login to connect your Microsoft 365 account.",
        });

        // Still wire up commands so /inbox_login works
        setCommandDependencies({ auth });
        return;
      }

      // 5. Start auto-refresh for Microsoft tokens
      auth.startAutoRefresh(async (error) => {
        log.error(`unified-inbox: token refresh failed: ${error}`);
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: `[Unified Inbox] Token refresh failed: ${error}\nUse /inbox_login to re-authenticate.`,
        });
      });

      // 6. Start monitors
      if (cfg.email.enabled) {
        emailMonitor = new EmailMonitor(cfg, auth, replyStore, log);
        await emailMonitor.start();
      }

      if (cfg.calendar.enabled) {
        calendarMonitor = new CalendarMonitor(cfg, auth, log);
        await calendarMonitor.start();
      }

      if (cfg.teamsChat.enabled) {
        teamsChatMonitor = new TeamsChatMonitor(cfg, auth, replyStore, log);
        await teamsChatMonitor.start();
      }

      // 7. Wire up command dependencies (with monitors)
      setCommandDependencies({
        auth,
        emailMonitor: emailMonitor ?? undefined,
        calendarMonitor: calendarMonitor ?? undefined,
        teamsChatMonitor: teamsChatMonitor ?? undefined,
      });

      // 8. Notify startup
      const monitors = [
        cfg.email.enabled ? "Email" : null,
        cfg.calendar.enabled ? "Calendar" : null,
        cfg.teamsChat.enabled ? "Teams Chat" : null,
        cfg.whatsapp.enabled ? "WhatsApp Bridge" : null,
      ]
        .filter(Boolean)
        .join(", ");

      log.info(`unified-inbox: all monitors started (${monitors})`);
      await sendTelegramMessage({
        botToken: cfg.telegramBotToken,
        chatId: cfg.telegramChatId,
        text: `[Unified Inbox] Started successfully. Active monitors: ${monitors}`,
      });
    },

    async stop(_ctx: OpenClawPluginServiceContext): Promise<void> {
      log.info("unified-inbox: service stopping...");

      emailMonitor?.stop();
      calendarMonitor?.stop();
      teamsChatMonitor?.stop();
      auth?.stopAutoRefresh();

      if (replyStore) {
        replyStore.stopAutoFlush();
        await replyStore.flush();
      }

      log.info("unified-inbox: service stopped");
    },
  };
}
