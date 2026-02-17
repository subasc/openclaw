// ============================================================================
// Email monitor: polls Microsoft Graph for new emails, forwards to Telegram
// ============================================================================

import type { ReplyStore } from "./reply-store.js";
import type { UnifiedInboxConfig } from "./config.js";
import type { MonitorStatus, EmailMessage, IMsAuthProvider } from "./types.js";
import { fetchMailDelta, markAsRead } from "./ms-graph-client.js";
import { sendTelegramMessage } from "./telegram-sender.js";
import { formatEmailNotification, formatEmailPlain } from "./formatters.js";
import { pushNotification } from "./notification-store.js";

type Logger = { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void };

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
    private log: Logger,
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

        // Process new messages
        const newMessages = result.messages.filter(
          (msg) => !this.seenIds.has(msg.id) && (!this.cfg.email.filterUnread || !msg.isRead),
        );

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
        this.log.error(
          "unified-inbox: email monitor paused after 3 consecutive failures",
        );
        await this.notifyError(
          "Email monitor paused after 3 failures. Last error: " + errorMsg,
        );
      }
    }
  }

  private async forwardToTelegram(email: EmailMessage): Promise<void> {
    const text = formatEmailNotification(email);

    const result = await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text,
      parseMode: "MarkdownV2",
    });

    // Push plain text to notification store for AI agent context
    pushNotification({
      source: "email",
      text: formatEmailPlain(email),
      timestamp: Date.now(),
    });

    if (result.ok && result.messageId) {
      // Store reply mapping so user can reply from Telegram
      this.replyStore.set(result.messageId, "email", {
        type: "email",
        messageId: email.id,
        conversationId: email.conversationId,
        fromAddress: email.from?.emailAddress?.address ?? "",
        subject: email.subject,
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
