// ============================================================================
// Unified Inbox service lifecycle
// Manages all monitors, token refresh, and graceful shutdown
// ============================================================================

import type { OpenClawPluginService, OpenClawPluginServiceContext } from "openclaw/plugin-sdk";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import type { UnifiedInboxConfig } from "./config.js";
import type { IMsAuthProvider } from "./types.js";
import type { ButtonContext } from "./types.js";
import { setToolAuthProviders } from "./agent-tools.js";
import { BrowserTokenAuth } from "./browser-token-auth.js";
import { CalendarMonitor } from "./calendar-monitor.js";
import { DirectReplyFlow } from "./direct-reply-flow.js";
import { EmailMonitor } from "./email-monitor.js";
import { ShortIdRegistry, EmailReplyFlow } from "./email-reply-flow.js";
import { getMe } from "./ms-graph-client.js";
import { setReplyRouterDependencies } from "./reply-router.js";
import { ReplyStore } from "./reply-store.js";
import { SubBotRegistry, createMonitorSubBot, createWhatsAppSubBot } from "./sub-bot-registry.js";
import { TeamsChatMonitor } from "./teams-chat-monitor.js";
import { setCommandDependencies } from "./telegram-commands.js";
import { sendTelegramMessage } from "./telegram-sender.js";
import { setWhatsAppBridgeReplyStore, setWhatsAppBridgeButtonRegistry } from "./whatsapp-bridge.js";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

export type UnifiedInboxServiceHandle = {
  service: OpenClawPluginService;
  getEmailReplyFlow: () => EmailReplyFlow | null;
  getDirectReplyFlow: () => DirectReplyFlow | null;
  getRegistry: () => SubBotRegistry | null;
  setWhatsAppSend: (fn: (jid: string, text: string) => Promise<void>) => void;
};

export function createUnifiedInboxService(
  cfg: UnifiedInboxConfig,
  log: Logger,
): UnifiedInboxServiceHandle {
  let auth: IMsAuthProvider | null = null;
  let teamsAuth: IMsAuthProvider | null = null;
  let outlookRestAuth: IMsAuthProvider | null = null;
  let replyStore: ReplyStore | null = null;
  let emailMonitor: EmailMonitor | null = null;
  let calendarMonitor: CalendarMonitor | null = null;
  let teamsChatMonitor: TeamsChatMonitor | null = null;
  const shortIdRegistry = new ShortIdRegistry<ButtonContext>();
  let emailReplyFlow: EmailReplyFlow | null = null;
  let directReplyFlow: DirectReplyFlow | null = null;
  let registry: SubBotRegistry | null = null;
  let pendingWhatsAppSend: ((jid: string, text: string) => Promise<void>) | null = null;

  const service: OpenClawPluginService = {
    id: "unified-inbox",

    async start(_ctx: OpenClawPluginServiceContext): Promise<void> {
      log.info("unified-inbox: service starting...");

      // 0. Bootstrap auto-approval config (never overwrites existing)
      const approvalFile = join(homedir(), ".openclaw", "exec-approvals.json");
      if (!existsSync(approvalFile)) {
        try {
          mkdirSync(dirname(approvalFile), { recursive: true });
          writeFileSync(
            approvalFile,
            JSON.stringify({ version: 1, defaults: { ask: "off" } }, null, 2) + "\n",
          );
          log.info(`unified-inbox: created auto-approval config at ${approvalFile}`);
        } catch (err) {
          log.warn(`unified-inbox: failed to create auto-approval config: ${String(err)}`);
        }
      }

      // 1. Initialize auth providers — CDP token extraction from live browser tabs
      const outlookAuth = new BrowserTokenAuth(
        { cdpPort: cfg.browserCdpPort, pageUrl: "https://outlook.office.com", pageName: "Outlook" },
        log,
      );
      auth = outlookAuth;

      if (cfg.teamsChat.enabled) {
        teamsAuth = new BrowserTokenAuth(
          {
            cdpPort: cfg.browserCdpPort,
            pageUrl: "https://teams.microsoft.com",
            pageName: "Teams",
          },
          log,
        );
      }

      // Outlook REST API auth (same tab, different token resource — has Mail.Send scope)
      outlookRestAuth = new BrowserTokenAuth(
        {
          cdpPort: cfg.browserCdpPort,
          pageUrl: "https://outlook.office.com",
          pageName: "Outlook-REST",
          resource: "outlook.office.com",
        },
        log,
      );

      log.info(
        "unified-inbox: using CDP token extraction (Outlook" + (teamsAuth ? " + Teams" : "") + ")",
      );

      // 2. Initialize reply store
      replyStore = new ReplyStore({
        storeFile: cfg.replyTracking.storeFile,
        maxEntries: cfg.replyTracking.maxEntries,
        ttlMs: cfg.replyTracking.ttlMs,
      });
      await replyStore.load();
      replyStore.startAutoFlush();

      // 3. Wire up shared dependencies (reply routing wired later once we know scopes)
      setWhatsAppBridgeReplyStore(replyStore);
      setWhatsAppBridgeButtonRegistry(shortIdRegistry);

      // 4. Try loading tokens from browser tabs (parallel — share the CDP wait)
      const [hasOutlookTokens, hasTeamsTokens, hasOutlookRestTokens] = await Promise.all([
        auth.loadPersistedTokens(),
        teamsAuth ? teamsAuth.loadPersistedTokens().catch(() => false) : Promise.resolve(false),
        outlookRestAuth.loadPersistedTokens().catch(() => false),
      ]);

      if (!hasOutlookTokens && !hasTeamsTokens) {
        log.info(
          "unified-inbox: no tokens found. Is the browser running and logged into Microsoft 365?",
        );
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: "[Unified Inbox] Started but no tokens found. Make sure the browser profile is running and you're logged into Microsoft 365.",
        });

        // Still wire up commands and reply router so /inbox_status works
        setReplyRouterDependencies({ auth, replyStore });
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

      if (hasOutlookRestTokens && outlookRestAuth) {
        outlookRestAuth.startAutoRefresh(async (error) => {
          log.error(`unified-inbox: Outlook REST token refresh failed: ${error}`);
        });
      }

      // 6. Start monitors with their respective auth providers.
      // Scopes are counter-intuitive: Outlook's Graph token has Chat.Read (not Mail.Read),
      // Teams' Graph token has Mail.Read/Calendars.Read (not Chat.Read).
      // This is because each web app uses its own Substrate API for its primary features
      // and Graph API only for cross-app features.
      const mailAuth = hasTeamsTokens && teamsAuth ? teamsAuth : auth;
      const chatAuth = hasOutlookTokens ? auth : teamsAuth;

      // Wire agent tools — use Outlook REST token for sending (has Mail.Send scope),
      // Graph token for reading (has Mail.Read scope), and
      // Teams token for calendar/tasks (has Calendars.ReadWrite + Tasks.ReadWrite scopes)
      const mailSendAuth = hasOutlookRestTokens && outlookRestAuth ? outlookRestAuth : mailAuth;
      setToolAuthProviders({
        mailAuth: mailSendAuth,
        mailReadAuth: mailAuth,
        calendarAuth: mailAuth, // Teams token — has Calendars.ReadWrite + Tasks.ReadWrite
        chatAuth: chatAuth!,
        log,
      });

      // Create EmailReplyFlow for interactive reply workflow
      emailReplyFlow = new EmailReplyFlow(cfg, shortIdRegistry, mailSendAuth, mailAuth, log);

      // Create DirectReplyFlow for Teams/WhatsApp inline button replies
      directReplyFlow = new DirectReplyFlow(
        cfg,
        shortIdRegistry,
        chatAuth ?? null,
        pendingWhatsAppSend, // may be set by index.ts before service starts
        log,
      );

      // Wire reply routing with scope-aware auth (after flows are created)
      setReplyRouterDependencies({
        auth: mailAuth, // email replies need Mail.Send (Teams token)
        teamsAuth: chatAuth, // Teams replies need Chat.ReadWrite (Outlook token)
        replyStore,
        emailReplyFlow,
        directReplyFlow,
      });

      if (cfg.email.enabled && mailAuth.isAuthenticated()) {
        // Auto-detect user email for smart filtering (To: field matching)
        let userEmail: string | undefined;
        try {
          const mailToken = await mailAuth.getAccessToken();
          const profile = await getMe(mailToken);
          userEmail = profile.mail;
          if (profile.displayName) {
            emailReplyFlow.setSenderName(profile.displayName);
          }
          log.info(`unified-inbox: user email detected: ${userEmail}`);
        } catch (err) {
          log.warn(
            `unified-inbox: failed to detect user email, filtering by To: disabled: ${String(err)}`,
          );
        }

        emailMonitor = new EmailMonitor(cfg, mailAuth, replyStore, shortIdRegistry, log, userEmail);
        await emailMonitor.start();
      }

      if (cfg.calendar.enabled && mailAuth.isAuthenticated()) {
        calendarMonitor = new CalendarMonitor(cfg, mailAuth, log);
        await calendarMonitor.start();
      }

      if (cfg.teamsChat.enabled && chatAuth?.isAuthenticated()) {
        teamsChatMonitor = new TeamsChatMonitor(cfg, chatAuth, replyStore, shortIdRegistry, log);
        await teamsChatMonitor.start();
      }

      // 7. Create sub-bot registry and register monitors
      registry = new SubBotRegistry(log);

      if (emailMonitor) {
        registry.register(
          createMonitorSubBot({
            id: "email",
            name: "Email Monitor",
            description: "Polls Microsoft Graph for new emails and forwards to Telegram",
            type: "email",
            monitor: emailMonitor,
          }),
        );
      }

      if (calendarMonitor) {
        registry.register(
          createMonitorSubBot({
            id: "calendar",
            name: "Calendar Monitor",
            description: "Polls Microsoft Graph for upcoming events and sends reminders",
            type: "calendar",
            monitor: calendarMonitor,
          }),
        );
      }

      if (teamsChatMonitor) {
        registry.register(
          createMonitorSubBot({
            id: "teams-chat",
            name: "Teams Chat Monitor",
            description: "Polls Microsoft Graph for new Teams chat messages",
            type: "teams-chat",
            monitor: teamsChatMonitor,
          }),
        );
      }

      if (cfg.whatsapp.enabled) {
        registry.register(createWhatsAppSubBot({}));
      }

      // 8. Wire up command dependencies (with monitors + registry)
      setCommandDependencies({
        auth,
        teamsAuth,
        outlookRestAuth,
        authMode: cfg.authMode,
        emailMonitor: emailMonitor ?? undefined,
        calendarMonitor: calendarMonitor ?? undefined,
        teamsChatMonitor: teamsChatMonitor ?? undefined,
        registry,
      });

      // 9. Wire auto-unpause: when tokens recover, unpause paused monitors
      const wireUnpause = (
        authProvider: IMsAuthProvider,
        authName: string,
        monitors: Array<{
          name: string;
          monitor: EmailMonitor | CalendarMonitor | TeamsChatMonitor | null;
        }>,
      ) => {
        if (typeof authProvider.setOnTokenRecovered !== "function") return;
        authProvider.setOnTokenRecovered(() => {
          const unpaused: string[] = [];
          for (const { name, monitor } of monitors) {
            if (monitor?.status?.paused) {
              monitor.status.paused = false;
              monitor.status.consecutiveFailures = 0;
              unpaused.push(name);
            }
          }
          if (unpaused.length > 0) {
            log.info(
              `unified-inbox: ${authName} token recovered — unpaused: ${unpaused.join(", ")}`,
            );
            sendTelegramMessage({
              botToken: cfg.telegramBotToken,
              chatId: cfg.telegramChatId,
              text: `[Unified Inbox] ${authName} token recovered. Unpaused monitors: ${unpaused.join(", ")}`,
            }).catch(() => {});
          }
        });
      };

      // mailAuth (Teams token) serves Email + Calendar monitors
      wireUnpause(mailAuth, "Mail/Calendar", [
        { name: "Email", monitor: emailMonitor },
        { name: "Calendar", monitor: calendarMonitor },
      ]);

      // chatAuth (Outlook token) serves Teams Chat monitor
      if (chatAuth) {
        wireUnpause(chatAuth, "Teams Chat", [{ name: "Teams Chat", monitor: teamsChatMonitor }]);
      }

      // 10. Notify startup with dashboard
      const allBots = registry.getAll();
      const statusLines = allBots.map((bot) => {
        const state = bot.status.paused ? "paused" : bot.status.running ? "running" : "stopped";
        return `${bot.name}: ${state}`;
      });

      const dashboardText = ["[Unified Inbox] Started", "", ...statusLines].join("\n");

      // Build inline buttons for each sub-bot (toggle start/stop)
      const botButtons = allBots.map((bot) => ({
        text: `${bot.status.running ? "Stop" : "Start"} ${bot.name.replace(" Monitor", "").replace(" Bridge", "")}`,
        callback_data: `inbox:sb:${bot.id}`,
      }));

      // Arrange buttons in rows of 2
      const buttonRows: Array<Array<{ text: string; callback_data: string }>> = [];
      for (let i = 0; i < botButtons.length; i += 2) {
        buttonRows.push(botButtons.slice(i, i + 2));
      }

      log.info(
        `unified-inbox: monitors started (${allBots.map((b) => b.name).join(", ") || "none"})`,
      );
      await sendTelegramMessage({
        botToken: cfg.telegramBotToken,
        chatId: cfg.telegramChatId,
        text: dashboardText,
        replyMarkup: buttonRows.length > 0 ? { inline_keyboard: buttonRows } : undefined,
      });
    },

    async stop(_ctx: OpenClawPluginServiceContext): Promise<void> {
      log.info("unified-inbox: service stopping...");

      emailMonitor?.stop();
      calendarMonitor?.stop();
      teamsChatMonitor?.stop();
      auth?.stopAutoRefresh();
      teamsAuth?.stopAutoRefresh();
      outlookRestAuth?.stopAutoRefresh();

      if (replyStore) {
        replyStore.stopAutoFlush();
        await replyStore.flush();
      }

      log.info("unified-inbox: service stopped");
    },
  };

  return {
    service,
    getEmailReplyFlow: () => emailReplyFlow,
    getDirectReplyFlow: () => directReplyFlow,
    getRegistry: () => registry,
    setWhatsAppSend: (fn: (jid: string, text: string) => Promise<void>) => {
      pendingWhatsAppSend = fn;
      // If DirectReplyFlow is already created, wire it immediately
      directReplyFlow?.setWhatsAppSend(fn);
    },
  };
}
