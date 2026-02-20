// ============================================================================
// Direct reply flow: lightweight state machine for Teams/WhatsApp replies
// (No LLM drafting — user types message, it's sent directly)
// ============================================================================

import type { UnifiedInboxConfig } from "./config.js";
import type { ShortIdRegistry } from "./email-reply-flow.js";
import type { ButtonContext, IMsAuthProvider } from "./types.js";
import { sendChatMessage } from "./ms-graph-client.js";
import { sendTelegramMessage, editTelegramMessage } from "./telegram-sender.js";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

type FlowState =
  | { phase: "idle" }
  | {
      phase: "awaiting_reply";
      shortId: string;
      source: "teams" | "whatsapp";
      startedAt: number;
    };

const TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes

export class DirectReplyFlow {
  private state: FlowState = { phase: "idle" };
  private suppressNextReply = false;
  private timeoutTimer: ReturnType<typeof setTimeout> | null = null;

  private whatsAppSend: ((jid: string, text: string) => Promise<void>) | null;

  constructor(
    private cfg: UnifiedInboxConfig,
    private registry: ShortIdRegistry<ButtonContext>,
    private teamsAuth: IMsAuthProvider | null,
    whatsAppSend: ((jid: string, text: string) => Promise<void>) | null,
    private log: Logger,
  ) {
    this.whatsAppSend = whatsAppSend;
  }

  /** Set WhatsApp send function (wired lazily after WhatsApp extension loads) */
  setWhatsAppSend(fn: (jid: string, text: string) => Promise<void>): void {
    this.whatsAppSend = fn;
  }

  // ---------- Public API for hooks ----------

  /**
   * Handle inbox:* callback strings for Teams/WhatsApp buttons.
   * Returns true if the content was handled.
   */
  async handleCallback(content: string): Promise<boolean> {
    if (!content.startsWith("inbox:")) return false;

    const parts = content.split(":");
    const action = parts[1];
    const shortId = parts[2];

    if (!action || !shortId) return false;

    switch (action) {
      case "tr": // Teams Reply
      case "tra": // Teams Reply All (same as Reply for Teams — goes to entire chat)
        return this.startReplyFlow(shortId, "teams");
      case "td": // Teams Delete/Dismiss
        return this.handleDismiss(shortId, "Teams");
      case "wr": // WhatsApp Reply
      case "wra": // WhatsApp Reply All (same as Reply — goes to entire chat)
        return this.startReplyFlow(shortId, "whatsapp");
      case "wd": // WhatsApp Delete/Dismiss
        return this.handleDismiss(shortId, "WhatsApp");
      default:
        return false;
    }
  }

  /**
   * Handle incoming user message when awaiting reply text.
   * Returns true if the message was consumed.
   */
  handleIncomingMessage(content: string): boolean {
    if (this.state.phase !== "awaiting_reply") return false;

    const text = content.trim();
    if (!text) return false;

    const { shortId, source } = this.state;

    // Transition synchronously
    this.state = { phase: "idle" };
    this.clearTimeout();
    this.suppressNextReply = true;

    // Fire-and-forget: send the message to the target platform
    this.sendReply(shortId, source, text).catch((err) => {
      this.log.error(
        `unified-inbox: direct reply send failed: ${err instanceof Error ? err.message : String(err)}`,
      );
    });

    return true;
  }

  /** Returns true (and clears flag) when a callback was just processed and the agent should stay silent. */
  shouldSuppressAgentReply(): boolean {
    if (this.suppressNextReply) {
      this.suppressNextReply = false;
      return true;
    }
    return false;
  }

  /** Whether the flow is busy (awaiting reply text). */
  isBusy(): boolean {
    return this.state.phase === "awaiting_reply";
  }

  // ---------- Private flow handlers ----------

  private async startReplyFlow(shortId: string, source: "teams" | "whatsapp"): Promise<boolean> {
    const ctx = this.registry.lookup(shortId);
    if (!ctx || ctx.type !== source) {
      this.log.warn(`unified-inbox: direct reply flow — unknown/mismatched shortId: ${shortId}`);
      return true; // Consumed but invalid
    }

    this.state = {
      phase: "awaiting_reply",
      shortId,
      source,
      startedAt: Date.now(),
    };
    this.suppressNextReply = true;

    const label = source === "teams" ? "Teams" : "WhatsApp";
    const target =
      ctx.type === "teams"
        ? `${ctx.senderName} in ${ctx.chatName}`
        : ctx.type === "whatsapp"
          ? `${ctx.senderName}${ctx.groupName ? ` in ${ctx.groupName}` : ""}`
          : "unknown";

    await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text: `Reply to ${target} on ${label} — Type your message:`,
    });

    this.startTimeout();
    return true;
  }

  private async handleDismiss(shortId: string, label: string): Promise<boolean> {
    const ctx = this.registry.lookup(shortId);
    if (!ctx) return true;

    this.suppressNextReply = true;

    const summary =
      ctx.type === "teams"
        ? `[${label}] ${ctx.senderName} in ${ctx.chatName}`
        : ctx.type === "whatsapp"
          ? `[${label}] ${ctx.senderName}${ctx.groupName ? ` in ${ctx.groupName}` : ""}`
          : `[${label}]`;

    await editTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      messageId: ctx.telegramMessageId,
      text: `${summary} (dismissed)`,
    });

    return true;
  }

  private async sendReply(
    shortId: string,
    source: "teams" | "whatsapp",
    text: string,
  ): Promise<void> {
    const ctx = this.registry.lookup(shortId);
    if (!ctx) return;

    try {
      if (source === "teams" && ctx.type === "teams") {
        if (!this.teamsAuth) throw new Error("Teams auth not available");
        const token = await this.teamsAuth.getAccessToken();
        await sendChatMessage(token, ctx.chatId, text);
        this.log.info(`unified-inbox: sent Teams reply to ${ctx.chatName}`);
      } else if (source === "whatsapp" && ctx.type === "whatsapp") {
        if (!this.whatsAppSend) throw new Error("WhatsApp send not available");
        const jid = ctx.groupJid || ctx.jid;
        await this.whatsAppSend(jid, text);
        this.log.info(`unified-inbox: sent WhatsApp reply to ${ctx.senderName}`);
      }

      await sendTelegramMessage({
        botToken: this.cfg.telegramBotToken,
        chatId: this.cfg.telegramChatId,
        text: `Sent to ${source === "teams" ? "Teams" : "WhatsApp"}`,
      });
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.log.error(`unified-inbox: direct reply failed: ${errorMsg}`);
      await sendTelegramMessage({
        botToken: this.cfg.telegramBotToken,
        chatId: this.cfg.telegramChatId,
        text: `Failed to send reply: ${errorMsg}`,
      });
    }
  }

  // ---------- Timeout / reset ----------

  private startTimeout(): void {
    this.clearTimeout();
    this.timeoutTimer = setTimeout(() => {
      if (this.state.phase !== "idle") {
        this.log.info("unified-inbox: direct reply flow timed out, resetting to idle");
        this.state = { phase: "idle" };
      }
    }, TIMEOUT_MS);
  }

  private clearTimeout(): void {
    if (this.timeoutTimer) {
      globalThis.clearTimeout(this.timeoutTimer);
      this.timeoutTimer = null;
    }
  }
}
