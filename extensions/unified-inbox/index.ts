import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import {
  listEmailsTool,
  sendEmailTool,
  replyEmailTool,
  listTeamsChatsTool,
  sendTeamsMessageTool,
  listCalendarTool,
  createEventTool,
  respondEventTool,
  deleteEventTool,
  listTasksTool,
  createTaskTool,
  completeTaskTool,
  updateTaskTool,
  deleteTaskTool,
} from "./src/agent-tools.js";
import { resolveUnifiedInboxConfig } from "./src/config.js";
import { formatNotificationsAsContext } from "./src/notification-store.js";
import { createReplyRouter, setReplyRouterWhatsAppSend } from "./src/reply-router.js";
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

    // Expose sub-bot status for Mission Control dashboard
    api.registerGatewayMethod("unified-inbox.status", ({ respond }) => {
      const registry = serviceHandle.getRegistry();
      if (!registry) {
        respond(true, { bots: [] });
        return;
      }
      const bots = registry.getAll().map((bot) => ({
        id: bot.id,
        name: bot.name,
        type: bot.type,
        status: {
          running: bot.status.running,
          paused: bot.status.paused,
          lastPollAt: bot.status.lastPollAt,
          lastErrorAt: bot.status.lastErrorAt,
          lastError: bot.status.lastError,
          consecutiveFailures: bot.status.consecutiveFailures,
        },
      }));
      respond(true, { bots });
    });

    // Register reply router hook (intercepts Telegram replies, routes back to source)
    const replyRouter = createReplyRouter(cfg, api.logger);
    api.on("message_received", replyRouter.handleMessageReceived);

    // Register WhatsApp bridge hook (forwards WhatsApp messages to Telegram)
    // Gated through the registry so /bot_stop whatsapp actually stops forwarding
    if (cfg.whatsapp.enabled) {
      const whatsappBridge = createWhatsAppBridge(cfg, api.logger);
      api.on("message_received", (event, ctx) => {
        const waBot = serviceHandle.getRegistry()?.get("whatsapp");
        if (waBot && !waBot.status.running) return;
        return whatsappBridge.handleMessageReceived(event, ctx);
      });

      // Wire WhatsApp send for reply router + direct reply flow
      try {
        const waSend = api.runtime.channel.whatsapp.sendMessageWhatsApp;
        const whatsAppSend = async (jid: string, text: string) => {
          await waSend(jid, text, { verbose: false });
        };
        setReplyRouterWhatsAppSend(whatsAppSend);
        serviceHandle.setWhatsAppSend(whatsAppSend);
      } catch {
        api.logger.warn(
          "unified-inbox: WhatsApp send function not available (WhatsApp extension may not be loaded)",
        );
      }
    }

    // Register Telegram slash commands
    registerInboxCommands(api, cfg);

    // Register agent tools — Email, Calendar, ToDo, Teams
    api.registerTool(listEmailsTool);
    api.registerTool(sendEmailTool);
    api.registerTool(replyEmailTool);
    api.registerTool(listCalendarTool);
    api.registerTool(createEventTool);
    api.registerTool(respondEventTool);
    api.registerTool(deleteEventTool);
    api.registerTool(listTasksTool);
    api.registerTool(createTaskTool);
    api.registerTool(completeTaskTool);
    api.registerTool(updateTaskTool);
    api.registerTool(deleteTaskTool);
    api.registerTool(listTeamsChatsTool);
    api.registerTool(sendTeamsMessageTool);

    // Handle inline button callbacks directly — bypasses agent pipeline entirely (no flash)
    api.on("telegram_callback", async (event) => {
      if (!event.data.startsWith("inbox:")) return;

      // Try email reply flow first
      const emailFlow = serviceHandle.getEmailReplyFlow();
      if (emailFlow) {
        const handled = await emailFlow.handleCallback(event.data);
        if (handled) return { handled: true };
      }

      // Try direct reply flow (Teams/WhatsApp)
      const directFlow = serviceHandle.getDirectReplyFlow();
      if (directFlow) {
        const handled = await directFlow.handleCallback(event.data);
        if (handled) return { handled: true };
      }

      // Handle sub-bot toggle buttons (inbox:sb:*)
      if (event.data.startsWith("inbox:sb:")) {
        const botId = event.data.slice("inbox:sb:".length);
        const registry = serviceHandle.getRegistry();
        if (registry) {
          const bot = registry.get(botId);
          if (bot) {
            let result: string;
            if (bot.status.running) {
              result = registry.stopBot(botId);
            } else {
              result = await registry.startBot(botId);
            }
            const { sendTelegramMessage: send } = await import("./src/telegram-sender.js");
            await send({
              botToken: cfg.telegramBotToken,
              chatId: cfg.telegramChatId,
              text: result,
            });
          }
        }
        return { handled: true };
      }
    });

    // Inject context into AI agent based on reply flow state
    api.on("before_agent_start", () => {
      const emailFlow = serviceHandle.getEmailReplyFlow();
      const directFlow = serviceHandle.getDirectReplyFlow();

      // If either flow is busy, suppress the agent.
      // Override systemPrompt to be minimal so the model generates NO_REPLY
      // as fast as possible — minimizes the brief streaming flash in Telegram.
      if (
        emailFlow?.shouldSuppressAgentReply() ||
        emailFlow?.isBusy() ||
        directFlow?.shouldSuppressAgentReply() ||
        directFlow?.isBusy()
      ) {
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
