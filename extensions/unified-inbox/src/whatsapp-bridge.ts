// ============================================================================
// WhatsApp -> Telegram bridge
// Hooks into message_received for WhatsApp, forwards to Telegram
// ============================================================================

import type { ReplyStore } from "./reply-store.js";
import type { UnifiedInboxConfig } from "./config.js";
import { sendTelegramMessage } from "./telegram-sender.js";
import { formatWhatsAppNotification } from "./formatters.js";

// Lazy-loaded shared reply store (set by service.ts)
let sharedReplyStore: ReplyStore | null = null;

export function setWhatsAppBridgeReplyStore(store: ReplyStore): void {
  sharedReplyStore = store;
}

type Logger = { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void };
type HookEvent = { from: string; content: string; timestamp?: number; metadata?: Record<string, unknown> };
type HookContext = { channelId: string; accountId?: string; conversationId?: string };

export function createWhatsAppBridge(
  cfg: UnifiedInboxConfig,
  log: Logger,
) {
  return {
    async handleMessageReceived(
      event: HookEvent,
      ctx: HookContext,
    ): Promise<void> {
      // Only process WhatsApp messages
      if (ctx.channelId !== "whatsapp") return;

      // Skip if bridge is disabled
      if (!cfg.whatsapp.enabled) return;

      // Skip empty messages
      if (!event.content?.trim()) return;

      const metadata = event.metadata ?? {};

      // Loop prevention: skip messages we sent ourselves
      if (metadata.isSelf || metadata.fromMe) return;

      const senderName =
        (metadata.senderName as string) ||
        (metadata.pushName as string) ||
        event.from ||
        "Unknown";
      const groupName = (metadata.groupName as string) || undefined;
      const groupJid = (metadata.groupJid as string) || undefined;
      const messageId = (metadata.messageId as string) || undefined;

      const text = formatWhatsAppNotification({
        from: event.from,
        content: event.content,
        senderName,
        groupName,
        groupJid,
        messageId,
      });

      const result = await sendTelegramMessage({
        botToken: cfg.telegramBotToken,
        chatId: cfg.telegramChatId,
        text,
        parseMode: "MarkdownV2",
      });

      if (result.ok && result.messageId && sharedReplyStore) {
        // The "from" for WhatsApp is typically the JID (e.g., "5511999999999@s.whatsapp.net")
        const jid = groupJid || event.from;
        sharedReplyStore.set(result.messageId, "whatsapp", {
          type: "whatsapp",
          jid,
          messageId,
          participant: groupJid ? event.from : undefined,
        });
      }

      if (!result.ok) {
        log.error(
          `unified-inbox: whatsapp bridge forward failed: ${result.error}`,
        );
      }
    },
  };
}
