// ============================================================================
// Email monitor: polls Microsoft Graph for new emails, forwards to Telegram
// ============================================================================

import type { UnifiedInboxConfig } from "./config.js";
import type { ShortIdRegistry } from "./email-reply-flow.js";
import type { ReplyStore } from "./reply-store.js";
import type { MonitorStatus, EmailMessage, IMsAuthProvider, ButtonContext } from "./types.js";
import { formatEmailWithSummary, formatEmailPlain } from "./formatters.js";
import { summarizeEmail } from "./llm-client.js";
import { fetchMailDelta, markAsRead } from "./ms-graph-client.js";
import { pushNotification } from "./notification-store.js";
import { sendTelegramMessage } from "./telegram-sender.js";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

export class EmailMonitor {
  private interval: ReturnType<typeof setInterval> | null = null;
  private deltaLinks: Map<string, string> = new Map();
  private seenIds: Set<string> = new Set();
  private shuttingDown = false;

  readonly status: MonitorStatus = {
    running: false,
    lastPollAt: null,
    lastErrorAt: null,
    lastError: null,
    consecutiveFailures: 0,
    paused: false,
  };

  constructor(
    private cfg: UnifiedInboxConfig,
    private auth: IMsAuthProvider,
    private replyStore: ReplyStore,
    private shortIdRegistry: ShortIdRegistry<ButtonContext>,
    private log: Logger,
    private userEmail?: string,
  ) {}

  async start(): Promise<void> {
    this.shuttingDown = false;
    this.status.running = true;
    this.status.paused = false;

    // Initial poll to establish baseline (mark existing emails as seen, don't flood)
    try {
      await this.initializeBaseline();
    } catch (err) {
      this.log.warn(`unified-inbox: email baseline failed: ${String(err)}`);
    }

    // Start polling loop
    this.interval = setInterval(async () => {
      if (this.shuttingDown || this.status.paused) return;
      await this.poll();
    }, this.cfg.email.pollIntervalMs);

    this.log.info(
      `unified-inbox: email monitor started (poll every ${this.cfg.email.pollIntervalMs / 1000}s)`,
    );
  }

  stop(): void {
    this.shuttingDown = true;
    this.status.running = false;
    if (this.interval) {
      clearInterval(this.interval);
      this.interval = null;
    }
  }

  /** Initialize delta links and mark current messages as seen */
  private async initializeBaseline(): Promise<void> {
    const token = await this.auth.getAccessToken();

    for (const folder of this.cfg.email.folders) {
      const result = await fetchMailDelta(token, folder, {
        filterUnread: this.cfg.email.filterUnread,
        top: 50,
      });

      // Store delta link for efficient subsequent polls
      if (result.deltaLink) {
        this.deltaLinks.set(folder, result.deltaLink);
      }

      // Mark all current messages as seen (don't notify on startup)
      for (const msg of result.messages) {
        this.seenIds.add(msg.id);
      }
    }
  }

  private async poll(): Promise<void> {
    try {
      const token = await this.auth.getAccessToken();

      for (const folder of this.cfg.email.folders) {
        const deltaLink = this.deltaLinks.get(folder) ?? folder;

        const result = await fetchMailDelta(token, deltaLink, {
          filterUnread: this.cfg.email.filterUnread,
          top: this.cfg.email.maxPerPoll,
        });

        // Update delta link
        if (result.deltaLink) {
          this.deltaLinks.set(folder, result.deltaLink);
        }

        // Process new messages — apply smart filtering
        const newMessages = result.messages
          .filter(
            (msg) => !this.seenIds.has(msg.id) && (!this.cfg.email.filterUnread || !msg.isRead),
          )
          .filter((msg) => this.shouldForward(msg));

        for (const msg of newMessages.slice(0, this.cfg.email.maxPerPoll)) {
          this.seenIds.add(msg.id);
          await this.forwardToTelegram(msg);
        }
      }

      this.status.lastPollAt = Date.now();
      this.status.consecutiveFailures = 0;

      // Trim seen set to prevent memory growth
      if (this.seenIds.size > 5000) {
        const arr = Array.from(this.seenIds);
        this.seenIds = new Set(arr.slice(arr.length - 2000));
      }
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.status.lastErrorAt = Date.now();
      this.status.lastError = errorMsg;
      this.status.consecutiveFailures++;
      this.log.error(`unified-inbox: email poll error: ${errorMsg}`);

      if (this.status.consecutiveFailures >= 3) {
        this.status.paused = true;
        this.log.error("unified-inbox: email monitor paused after 3 consecutive failures");
        await this.notifyError("Email monitor paused after 3 failures. Last error: " + errorMsg);
      }
    }
  }

  private shouldForward(email: EmailMessage): boolean {
    const filter = this.cfg.email.filter;
    if (!filter.enabled) return true;

    // 1. User is a direct recipient (To: field)
    if (this.userEmail && email.toRecipients?.length) {
      const userLower = this.userEmail.toLowerCase();
      const isDirectRecipient = email.toRecipients.some(
        (r) => r.emailAddress?.address?.toLowerCase() === userLower,
      );
      if (isDirectRecipient) return true;
    }

    // 2. Email has high importance (urgent flag)
    if (email.importance === "high") return true;

    // 3. Email body mentions configured keywords (e.g. @Subas)
    if (filter.bodyMentionKeywords.length > 0) {
      const bodyText = (
        (email.body?.content ?? "") +
        " " +
        (email.bodyPreview ?? "")
      ).toLowerCase();
      const hasMention = filter.bodyMentionKeywords.some((kw) =>
        bodyText.includes(kw.toLowerCase()),
      );
      if (hasMention) return true;
    }

    return false;
  }

  private async forwardToTelegram(email: EmailMessage): Promise<void> {
    const from = email.from?.emailAddress;
    const fromName = from?.name || from?.address || "Unknown";
    const fromAddr = from?.address || "";

    // LLM summarization (falls back to bodyPreview on failure)
    const emailBody = email.body?.content
      ? email.body.content
          .replace(/<[^>]+>/g, "")
          .replace(/&nbsp;/g, " ")
          .trim()
      : email.bodyPreview || "";
    const summary = await summarizeEmail(
      { from: `${fromName} <${fromAddr}>`, subject: email.subject, body: emailBody },
      this.log,
    );

    const text = formatEmailWithSummary(email, summary);

    // Register in ShortIdRegistry for inline button callbacks
    // We need the telegramMessageId after sending, so register with a placeholder first
    const shortId = this.shortIdRegistry.register({
      type: "email",
      messageId: email.id,
      fromName,
      fromAddress: fromAddr,
      subject: email.subject,
      bodyPreview: email.bodyPreview || "",
      toRecipients: (email.toRecipients || []).map((r) => ({
        name: r.emailAddress?.name || "",
        address: r.emailAddress?.address || "",
      })),
      telegramMessageId: 0, // Updated after send
      registeredAt: Date.now(),
    });

    const result = await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text,
      parseMode: "MarkdownV2",
      replyMarkup: {
        inline_keyboard: [
          [
            { text: "Reply", callback_data: `inbox:r:${shortId}` },
            { text: "Reply All", callback_data: `inbox:ra:${shortId}` },
            { text: "Delete", callback_data: `inbox:d:${shortId}` },
          ],
        ],
      },
    });

    // Push plain text to notification store for AI agent context
    pushNotification({
      source: "email",
      text: formatEmailPlain(email),
      timestamp: Date.now(),
    });

    if (result.ok && result.messageId) {
      // Update the ShortIdRegistry entry with the actual Telegram message ID
      const ctx = this.shortIdRegistry.lookup(shortId);
      if (ctx) ctx.telegramMessageId = result.messageId;

      // Store reply mapping so user can reply from Telegram (existing reply-to-message flow)
      this.replyStore.set(result.messageId, "email", {
        type: "email",
        messageId: email.id,
        conversationId: email.conversationId,
        fromAddress: fromAddr,
        subject: email.subject,
        fromName,
        bodyPreview: email.bodyPreview || "",
        toRecipients: (email.toRecipients || []).map((r) => ({
          name: r.emailAddress?.name || "",
          address: r.emailAddress?.address || "",
        })),
      });
    }
  }

  private async notifyError(message: string): Promise<void> {
    await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text: `[Unified Inbox] ${message}`,
    });
  }
}
