// ============================================================================
// Email reply flow: ShortIdRegistry + state machine for interactive replies
// ============================================================================

import type { UnifiedInboxConfig } from "./config.js";
import type { ButtonContext, IMsAuthProvider } from "./types.js";
import { formatDraftPreview } from "./formatters.js";
import { draftReply } from "./llm-client.js";
import {
  replyToEmailViaOutlookRest,
  replyAllToEmailViaOutlookRest,
  markAsRead,
} from "./ms-graph-client.js";
import { sendTelegramMessage, editTelegramMessage } from "./telegram-sender.js";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

// ============================================================================
// ShortIdRegistry — maps compact base-36 IDs to full email contexts
// Telegram limits callback_data to 64 bytes; Graph message IDs are ~150 chars
// ============================================================================

export type EmailContext = {
  type: "email";
  messageId: string;
  fromName: string;
  fromAddress: string;
  subject: string;
  bodyPreview: string;
  toRecipients: Array<{ name: string; address: string }>;
  telegramMessageId: number;
  registeredAt: number;
};

export class ShortIdRegistry<T extends { registeredAt: number } = EmailContext> {
  private counter = 0;
  private map = new Map<string, T>();
  private static readonly MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24h

  register(context: T): string {
    this.prune();
    const id = this.counter.toString(36);
    this.counter++;
    this.map.set(id, context);
    return id;
  }

  lookup(shortId: string): T | undefined {
    const ctx = this.map.get(shortId);
    if (!ctx) return undefined;
    if (Date.now() - ctx.registeredAt > ShortIdRegistry.MAX_AGE_MS) {
      this.map.delete(shortId);
      return undefined;
    }
    return ctx;
  }

  private prune(): void {
    const now = Date.now();
    for (const [id, ctx] of this.map) {
      if (now - ctx.registeredAt > ShortIdRegistry.MAX_AGE_MS) {
        this.map.delete(id);
      }
    }
  }
}

// ============================================================================
// EmailReplyFlow — state machine for the interactive reply workflow
//
// States: idle → awaiting_notes → drafting → awaiting_approval → idle
// ============================================================================

type FlowState =
  | { phase: "idle" }
  | { phase: "awaiting_notes"; shortId: string; mode: "reply" | "reply-all"; startedAt: number }
  | { phase: "drafting"; shortId: string; mode: "reply" | "reply-all"; notes: string }
  | {
      phase: "awaiting_approval";
      shortId: string;
      mode: "reply" | "reply-all";
      draft: string;
      draftMessageId?: number;
    };

const TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes

export class EmailReplyFlow {
  private state: FlowState = { phase: "idle" };
  private suppressNextReply = false;
  private timeoutTimer: ReturnType<typeof setTimeout> | null = null;

  private senderName = "Subas";

  constructor(
    private cfg: UnifiedInboxConfig,
    private registry: ShortIdRegistry<ButtonContext>,
    private mailSendAuth: IMsAuthProvider,
    private mailReadAuth: IMsAuthProvider,
    private log: Logger,
  ) {}

  /** Narrow a ButtonContext lookup to EmailContext */
  private lookupEmail(shortId: string): EmailContext | undefined {
    const ctx = this.lookupEmail(shortId);
    if (!ctx || ctx.type !== "email") return undefined;
    return ctx as EmailContext;
  }

  /** Set the display name used for email sign-offs (e.g. "Regards, Subas") */
  setSenderName(name: string): void {
    this.senderName = name;
  }

  // ---------- Public API for hooks ----------

  /**
   * Handle inbox:* callback strings (from Telegram inline buttons).
   * Returns true if the content was handled (caller should not process further).
   */
  async handleCallback(content: string): Promise<boolean> {
    if (!content.startsWith("inbox:")) return false;

    const parts = content.split(":");
    const action = parts[1];
    const shortId = parts[2];

    if (!action || !shortId) return false;

    switch (action) {
      case "r": // Reply
      case "ra": // Reply All
        return this.startReplyFlow(shortId, action === "ra" ? "reply-all" : "reply");
      case "d": // Delete / dismiss
        return this.handleDelete(shortId);
      case "as": // Approve & Send
        return this.handleApproveSend();
      case "dd": // Discard draft
        return this.handleDiscard();
      default:
        return false;
    }
  }

  /**
   * Handle incoming user message when awaiting notes.
   * Sets state synchronously, then kicks off async draft generation directly.
   * Returns true if the message was consumed as reply notes.
   */
  handleIncomingMessage(content: string): boolean {
    if (this.state.phase !== "awaiting_notes") return false;

    const notes = content.trim();
    if (!notes) return false;

    // Capture values before state transition
    const shortId = this.state.shortId;
    const mode = this.state.mode;

    // Transition synchronously (critical: message_received is fire-and-forget)
    this.state = {
      phase: "drafting",
      shortId,
      mode,
      notes,
    };
    this.clearTimeout();
    this.suppressNextReply = true;

    // Fire-and-forget: generate draft via direct LLM call and show with buttons
    this.generateAndShowDraft(shortId, mode, notes).catch((err) => {
      this.log.error(
        `unified-inbox: draft generation failed: ${err instanceof Error ? err.message : String(err)}`,
      );
      this.reset();
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

  /** Whether the flow is busy (drafting or awaiting approval) and the agent should stay silent. */
  isBusy(): boolean {
    return this.state.phase === "drafting" || this.state.phase === "awaiting_approval";
  }

  /** Generate draft via direct LLM call and show with approval buttons. */
  private async generateAndShowDraft(
    shortId: string,
    mode: "reply" | "reply-all",
    notes: string,
  ): Promise<void> {
    const ctx = this.lookupEmail(shortId);
    if (!ctx) {
      this.reset();
      return;
    }

    // Send "Drafting..." status message
    await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text: "Drafting reply...",
    });

    // Call LLM directly to generate the draft
    const draft = await draftReply(
      {
        fromName: ctx.fromName,
        fromAddress: ctx.fromAddress,
        subject: ctx.subject,
        bodyPreview: ctx.bodyPreview,
        mode,
        notes,
        senderName: this.senderName,
      },
      this.log,
    );

    if (!draft) {
      await sendTelegramMessage({
        botToken: this.cfg.telegramBotToken,
        chatId: this.cfg.telegramChatId,
        text: "Failed to generate draft. Please try again.",
      });
      this.reset();
      return;
    }

    // Show draft with approval buttons
    const draftText = formatDraftPreview({
      fromName: ctx.fromName,
      fromAddress: ctx.fromAddress,
      subject: ctx.subject,
      mode,
      draft,
    });

    const result = await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text: draftText,
      parseMode: "MarkdownV2",
      replyMarkup: {
        inline_keyboard: [
          [
            { text: "Approve & Send", callback_data: `inbox:as:${shortId}` },
            { text: "Discard", callback_data: `inbox:dd:${shortId}` },
          ],
        ],
      },
    });

    this.state = {
      phase: "awaiting_approval",
      shortId,
      mode,
      draft,
      draftMessageId: result.messageId,
    };

    this.startTimeout();
  }

  // ---------- Private flow handlers ----------

  private async startReplyFlow(shortId: string, mode: "reply" | "reply-all"): Promise<boolean> {
    const ctx = this.lookupEmail(shortId);
    if (!ctx) {
      this.log.warn(`unified-inbox: reply flow — unknown shortId: ${shortId}`);
      return true; // Consumed but invalid
    }

    this.state = {
      phase: "awaiting_notes",
      shortId,
      mode,
      startedAt: Date.now(),
    };
    this.suppressNextReply = true;

    const modeLabel = mode === "reply-all" ? "Reply All" : "Reply";
    await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text: `${modeLabel} to ${ctx.fromName} — "${ctx.subject}"\n\nType your notes (bullet points, shorthand — I'll draft a professional reply):`,
    });

    this.startTimeout();
    return true;
  }

  private async handleDelete(shortId: string): Promise<boolean> {
    const ctx = this.lookupEmail(shortId);
    if (!ctx) return true;

    this.suppressNextReply = true;

    try {
      const token = await this.mailReadAuth.getAccessToken();
      await markAsRead(token, ctx.messageId);
    } catch (err) {
      this.log.error(
        `unified-inbox: markAsRead failed: ${err instanceof Error ? err.message : String(err)}`,
      );
    }

    // Edit notification to remove buttons
    await editTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      messageId: ctx.telegramMessageId,
      text: `[Email] from ${ctx.fromName} — "${ctx.subject}" (dismissed)`,
    });

    return true;
  }

  private async handleApproveSend(): Promise<boolean> {
    if (this.state.phase !== "awaiting_approval") return true;

    const ctx = this.lookupEmail(this.state.shortId);
    if (!ctx) {
      this.reset();
      return true;
    }

    this.suppressNextReply = true;
    const { draft, mode, draftMessageId } = this.state;

    try {
      const token = await this.mailSendAuth.getAccessToken();
      if (mode === "reply-all") {
        await replyAllToEmailViaOutlookRest(token, ctx.messageId, draft);
      } else {
        await replyToEmailViaOutlookRest(token, ctx.messageId, draft);
      }

      this.log.info(`unified-inbox: sent ${mode} to ${ctx.fromAddress} (${ctx.subject})`);

      // Edit draft message to show "Sent"
      if (draftMessageId) {
        await editTelegramMessage({
          botToken: this.cfg.telegramBotToken,
          chatId: this.cfg.telegramChatId,
          messageId: draftMessageId,
          text: `Sent ${mode === "reply-all" ? "reply-all" : "reply"} to ${ctx.fromName} — "${ctx.subject}"`,
        });
      }
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.log.error(`unified-inbox: send failed: ${errorMsg}`);
      await sendTelegramMessage({
        botToken: this.cfg.telegramBotToken,
        chatId: this.cfg.telegramChatId,
        text: `Failed to send reply: ${errorMsg}`,
      });
    }

    this.reset();
    return true;
  }

  private async handleDiscard(): Promise<boolean> {
    if (this.state.phase !== "awaiting_approval") return true;

    this.suppressNextReply = true;
    const { draftMessageId } = this.state;

    if (draftMessageId) {
      await editTelegramMessage({
        botToken: this.cfg.telegramBotToken,
        chatId: this.cfg.telegramChatId,
        messageId: draftMessageId,
        text: "Draft discarded.",
      });
    }

    this.reset();
    return true;
  }

  // ---------- Timeout / reset ----------

  private startTimeout(): void {
    this.clearTimeout();
    this.timeoutTimer = setTimeout(() => {
      if (this.state.phase !== "idle") {
        this.log.info("unified-inbox: reply flow timed out, resetting to idle");
        this.reset();
      }
    }, TIMEOUT_MS);
  }

  private clearTimeout(): void {
    if (this.timeoutTimer) {
      globalThis.clearTimeout(this.timeoutTimer);
      this.timeoutTimer = null;
    }
  }

  private reset(): void {
    this.state = { phase: "idle" };
    this.clearTimeout();
  }
}
