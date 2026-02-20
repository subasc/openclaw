// ============================================================================
// Reply router: intercepts Telegram replies and routes them back to source
// ============================================================================

import type { UnifiedInboxConfig } from "./config.js";
import type { DirectReplyFlow } from "./direct-reply-flow.js";
import type { EmailReplyFlow } from "./email-reply-flow.js";
import type { ReplyStore } from "./reply-store.js";
import type { IMsAuthProvider } from "./types.js";
import { replyToEmail, sendChatMessage } from "./ms-graph-client.js";
import { sendTelegramMessage } from "./telegram-sender.js";

// Shared state set by service.ts after initialization
let sharedAuth: IMsAuthProvider | null = null;
let sharedTeamsAuth: IMsAuthProvider | null = null;
let sharedReplyStore: ReplyStore | null = null;
let sharedWhatsAppSend: ((jid: string, text: string) => Promise<void>) | null = null;
let sharedEmailReplyFlow: EmailReplyFlow | null = null;
let sharedDirectReplyFlow: DirectReplyFlow | null = null;

export function setReplyRouterDependencies(deps: {
  auth: IMsAuthProvider;
  teamsAuth?: IMsAuthProvider;
  replyStore: ReplyStore;
  whatsAppSend?: (jid: string, text: string) => Promise<void>;
  emailReplyFlow?: EmailReplyFlow;
  directReplyFlow?: DirectReplyFlow;
}): void {
  sharedAuth = deps.auth;
  sharedTeamsAuth = deps.teamsAuth ?? deps.auth;
  sharedReplyStore = deps.replyStore;
  if (deps.whatsAppSend) sharedWhatsAppSend = deps.whatsAppSend;
  if (deps.emailReplyFlow) sharedEmailReplyFlow = deps.emailReplyFlow;
  if (deps.directReplyFlow) sharedDirectReplyFlow = deps.directReplyFlow;
}

/** Wire WhatsApp send function (called from index.ts after WhatsApp extension loads) */
export function setReplyRouterWhatsAppSend(fn: (jid: string, text: string) => Promise<void>): void {
  sharedWhatsAppSend = fn;
}

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};
type HookEvent = {
  from: string;
  content: string;
  timestamp?: number;
  metadata?: Record<string, unknown>;
};
type HookContext = { channelId: string; accountId?: string; conversationId?: string };

export function createReplyRouter(cfg: UnifiedInboxConfig, log: Logger) {
  return {
    async handleMessageReceived(event: HookEvent, ctx: HookContext): Promise<void> {
      // Only process Telegram messages
      if (ctx.channelId !== "telegram") return;

      // Capture user's reply notes when in awaiting_notes/awaiting_reply state.
      // Don't return — let the message continue to the agent pipeline.
      // (inbox:* callbacks are now handled via the telegram_callback hook,
      //  which bypasses the agent pipeline entirely — no "No" flash.)
      if (sharedEmailReplyFlow) {
        const consumed = sharedEmailReplyFlow.handleIncomingMessage(event.content ?? "");
        if (consumed) return;
      }
      if (sharedDirectReplyFlow) {
        const consumed = sharedDirectReplyFlow.handleIncomingMessage(event.content ?? "");
        if (consumed) return;
      }

      // Check if this is a reply to a forwarded message
      const replyToId = (event.metadata?.replyToMessageId as number) ?? null;
      if (!replyToId || !sharedReplyStore) return;

      const mapping = sharedReplyStore.get(replyToId);
      if (!mapping) return;

      const replyText = event.content?.trim();
      if (!replyText) return;

      try {
        switch (mapping.sourceContext.type) {
          case "email":
            await routeToEmail(replyText, mapping.sourceContext.messageId, log);
            break;

          case "teams-chat":
            await routeToTeams(replyText, mapping.sourceContext.chatId, log);
            break;

          case "whatsapp":
            await routeToWhatsApp(replyText, mapping.sourceContext.jid, log);
            break;
        }

        // Confirm delivery with a small reaction-like message
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: `Sent to ${mapping.source}`,
        });
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        log.error(`unified-inbox: reply routing failed (${mapping.source}): ${errorMsg}`);
        await sendTelegramMessage({
          botToken: cfg.telegramBotToken,
          chatId: cfg.telegramChatId,
          text: `Failed to send reply to ${mapping.source}: ${errorMsg}`,
        });
      }
    },
  };
}

async function routeToEmail(text: string, messageId: string, log: Logger): Promise<void> {
  if (!sharedAuth) throw new Error("Not authenticated");
  const token = await sharedAuth.getAccessToken();
  await replyToEmail(token, messageId, text);
  log.info(`unified-inbox: replied to email ${messageId}`);
}

async function routeToTeams(text: string, chatId: string, log: Logger): Promise<void> {
  if (!sharedTeamsAuth) throw new Error("Teams auth not available");
  const token = await sharedTeamsAuth.getAccessToken();
  await sendChatMessage(token, chatId, text);
  log.info(`unified-inbox: sent Teams message to chat ${chatId}`);
}

async function routeToWhatsApp(text: string, jid: string, log: Logger): Promise<void> {
  if (!sharedWhatsAppSend) {
    throw new Error("WhatsApp send not available");
  }
  await sharedWhatsAppSend(jid, text);
  log.info(`unified-inbox: sent WhatsApp message to ${jid}`);
}
