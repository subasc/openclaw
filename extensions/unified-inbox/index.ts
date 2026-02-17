import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import { resolveUnifiedInboxConfig } from "./src/config.js";
import { createUnifiedInboxService } from "./src/service.js";
import { createReplyRouter } from "./src/reply-router.js";
import { createWhatsAppBridge } from "./src/whatsapp-bridge.js";
import { registerInboxCommands } from "./src/telegram-commands.js";
import { formatNotificationsAsContext } from "./src/notification-store.js";

const plugin = {
  id: "unified-inbox",
  name: "Unified Inbox",
  description:
    "Bridge Microsoft 365 email, calendar, Teams chat, and WhatsApp through Telegram",
  configSchema: {
    safeParse(value: unknown) {
      return { success: true, data: value };
    },
    jsonSchema: {
      type: "object",
      additionalProperties: true,
      properties: {},
    },
  },

  register(api: OpenClawPluginApi) {
    const cfg = resolveUnifiedInboxConfig(api.pluginConfig);
    if (!cfg.enabled) {
      api.logger.info("unified-inbox: disabled in config");
      return;
    }

    if (!cfg.telegramChatId || !cfg.telegramBotToken) {
      api.logger.warn(
        "unified-inbox: telegramChatId and telegramBotToken are required",
      );
      return;
    }

    // Register the background service (manages all monitors + token refresh)
    api.registerService(createUnifiedInboxService(cfg, api.logger));

    // Register reply router hook (intercepts Telegram replies, routes back to source)
    const replyRouter = createReplyRouter(cfg, api.logger);
    api.on("message_received", replyRouter.handleMessageReceived);

    // Register WhatsApp bridge hook (forwards WhatsApp messages to Telegram)
    if (cfg.whatsapp.enabled) {
      const whatsappBridge = createWhatsAppBridge(cfg, api.logger);
      api.on("message_received", whatsappBridge.handleMessageReceived);
    }

    // Register Telegram slash commands
    registerInboxCommands(api, cfg);

    // Inject recent notifications into AI agent context so it can answer
    // questions about emails, calendar events, and Teams messages
    api.on("before_agent_start", () => {
      const context = formatNotificationsAsContext();
      if (context) {
        return { prependContext: context };
      }
    });

    api.logger.info("unified-inbox: registered successfully");
  },
};

export default plugin;
