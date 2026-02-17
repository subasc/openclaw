import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import {
  sendEmailTool,
  replyEmailTool,
  listTeamsChatsTool,
  sendTeamsMessageTool,
} from "./src/agent-tools.js";
import { resolveUnifiedInboxConfig } from "./src/config.js";
import { formatNotificationsAsContext } from "./src/notification-store.js";
import { createReplyRouter } from "./src/reply-router.js";
import { createUnifiedInboxService } from "./src/service.js";
import { registerInboxCommands } from "./src/telegram-commands.js";
import { createWhatsAppBridge } from "./src/whatsapp-bridge.js";

const plugin = {
  id: "unified-inbox",
  name: "Unified Inbox",
  description: "Bridge Microsoft 365 email, calendar, Teams chat, and WhatsApp through Telegram",
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
      api.logger.warn("unified-inbox: telegramChatId and telegramBotToken are required");
      return;
    }

    // Register the background service (manages all monitors + token refresh)
    const serviceHandle = createUnifiedInboxService(cfg, api.logger);
    api.registerService(serviceHandle.service);

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

    // Register agent tools so the AI can send emails and Teams messages
    api.registerTool(sendEmailTool);
    api.registerTool(replyEmailTool);
    api.registerTool(listTeamsChatsTool);
    api.registerTool(sendTeamsMessageTool);

    // Handle inline button callbacks directly — bypasses agent pipeline entirely (no flash)
    api.on("telegram_callback", async (event) => {
      const flow = serviceHandle.getEmailReplyFlow();
      if (!flow || !event.data.startsWith("inbox:")) return;
      await flow.handleCallback(event.data);
      return { handled: true };
    });

    // Inject context into AI agent based on reply flow state
    api.on("before_agent_start", () => {
      const flow = serviceHandle.getEmailReplyFlow();

      // If the flow is busy (drafting or awaiting approval), suppress the agent.
      // Override systemPrompt to be minimal so the model generates NO_REPLY
      // as fast as possible — minimizes the brief streaming flash in Telegram.
      if (flow?.shouldSuppressAgentReply() || flow?.isBusy()) {
        return {
          systemPrompt: "Output exactly: NO_REPLY",
          prependContext: "NO_REPLY",
        };
      }

      // Default: inject recent notifications for AI context
      const context = formatNotificationsAsContext();
      if (context) {
        return { prependContext: context };
      }
    });

    api.logger.info("unified-inbox: registered successfully");
  },
};

export default plugin;
