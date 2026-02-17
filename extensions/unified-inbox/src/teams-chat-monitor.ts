// ============================================================================
// Teams chat monitor: polls Microsoft Graph for new chat messages
// ============================================================================

import type { ReplyStore } from "./reply-store.js";
import type { UnifiedInboxConfig } from "./config.js";
import type { MonitorStatus, TeamsChat, TeamsChatMessage, IMsAuthProvider } from "./types.js";
import { listChats, listChatMessages } from "./ms-graph-client.js";
import { sendTelegramMessage } from "./telegram-sender.js";
import { formatTeamsChatNotification, formatTeamsChatPlain } from "./formatters.js";
import { pushNotification } from "./notification-store.js";

type Logger = { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void };

export class TeamsChatMonitor {
  private interval: ReturnType<typeof setInterval> | null = null;
  private lastMessageTimestamps: Map<string, string> = new Map(); // chatId -> ISO timestamp
  private myUserId: string | null = null;
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

    // Initialize: get current user ID and snapshot current message timestamps
    try {
      await this.initializeBaseline();
    } catch (err) {
      this.log.warn(
        `unified-inbox: teams chat baseline failed: ${String(err)}`,
      );
    }

    this.interval = setInterval(async () => {
      if (this.shuttingDown || this.status.paused) return;
      await this.poll();
    }, this.cfg.teamsChat.pollIntervalMs);

    this.log.info(
      `unified-inbox: teams chat monitor started (poll every ${this.cfg.teamsChat.pollIntervalMs / 1000}s)`,
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

  /** Fetch recent chats for on-demand /teams command */
  async fetchRecentChats(): Promise<TeamsChat[]> {
    const token = await this.auth.getAccessToken();
    return listChats(token);
  }

  private async initializeBaseline(): Promise<void> {
    const token = await this.auth.getAccessToken();

    // Get current user ID to filter out own messages
    try {
      const { getMe } = await import("./ms-graph-client.js");
      const me = await getMe(token);
      this.myUserId = me.id;
    } catch {
      // Non-critical
    }

    // Snapshot current chat timestamps
    const chats = await listChats(token);
    for (const chat of chats) {
      if (chat.lastMessagePreview?.createdDateTime) {
        this.lastMessageTimestamps.set(
          chat.id,
          chat.lastMessagePreview.createdDateTime,
        );
      }
    }
  }

  private async poll(): Promise<void> {
    try {
      const token = await this.auth.getAccessToken();
      const chats = await listChats(token);

      for (const chat of chats) {
        const lastPreview = chat.lastMessagePreview;
        if (!lastPreview?.createdDateTime) continue;

        const previousTimestamp = this.lastMessageTimestamps.get(chat.id);

        // Check if there's a new message since last check
        if (
          previousTimestamp &&
          lastPreview.createdDateTime > previousTimestamp
        ) {
          // Fetch new messages
          const messages = await listChatMessages(token, chat.id, {
            since: previousTimestamp,
            top: 10,
          });

          for (const msg of messages) {
            // Skip system messages
            if (msg.messageType !== "message") continue;

            // Skip own messages
            if (this.myUserId && msg.from?.user?.id === this.myUserId) continue;

            // Skip bot messages if configured
            if (this.cfg.teamsChat.excludeBotMessages && msg.from?.application) {
              continue;
            }

            await this.forwardToTelegram(msg, chat);
          }
        }

        // Update timestamp
        this.lastMessageTimestamps.set(
          chat.id,
          lastPreview.createdDateTime,
        );
      }

      this.status.lastPollAt = Date.now();
      this.status.consecutiveFailures = 0;
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.status.lastErrorAt = Date.now();
      this.status.lastError = errorMsg;
      this.status.consecutiveFailures++;
      this.log.error(`unified-inbox: teams chat poll error: ${errorMsg}`);

      if (this.status.consecutiveFailures >= 3) {
        this.status.paused = true;
        this.log.error(
          "unified-inbox: teams chat monitor paused after 3 consecutive failures",
        );
        await sendTelegramMessage({
          botToken: this.cfg.telegramBotToken,
          chatId: this.cfg.telegramChatId,
          text: `[Unified Inbox] Teams chat monitor paused after 3 failures. Last error: ${errorMsg}`,
        });
      }
    }
  }

  private async forwardToTelegram(
    msg: TeamsChatMessage,
    chat: TeamsChat,
  ): Promise<void> {
    const chatName =
      chat.topic ||
      (chat.chatType === "oneOnOne"
        ? msg.from?.user?.displayName || "Direct Message"
        : "Group Chat");

    const text = formatTeamsChatNotification(msg, chatName);

    const result = await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text,
      parseMode: "MarkdownV2",
    });

    // Push plain text to notification store for AI agent context
    pushNotification({
      source: "teams",
      text: formatTeamsChatPlain(msg, chatName),
      timestamp: Date.now(),
    });

    if (result.ok && result.messageId) {
      this.replyStore.set(result.messageId, "teams-chat", {
        type: "teams-chat",
        chatId: chat.id,
        messageId: msg.id,
      });
    }
  }
}
