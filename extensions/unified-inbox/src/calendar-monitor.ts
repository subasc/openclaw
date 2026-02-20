// ============================================================================
// Calendar monitor: polls Microsoft Graph for upcoming events, sends reminders
// ============================================================================

import type { UnifiedInboxConfig } from "./config.js";
import type { MonitorStatus, CalendarEvent, IMsAuthProvider } from "./types.js";
import { fetchCalendarView } from "./ms-graph-client.js";
import { sendTelegramMessage } from "./telegram-sender.js";
import {
  formatCalendarNotification,
  formatCalendarReminder,
  formatCalendarReminderPlain,
} from "./formatters.js";
import { pushNotification } from "./notification-store.js";

type Logger = { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void };

export class CalendarMonitor {
  private interval: ReturnType<typeof setInterval> | null = null;
  private notifiedEvents: Map<string, Set<number>> = new Map(); // eventId -> set of reminderMinutes already sent
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
    private log: Logger,
  ) {}

  async start(): Promise<void> {
    this.shuttingDown = false;
    this.status.running = true;
    this.status.paused = false;

    // Initial poll
    await this.poll().catch((err) =>
      this.log.warn(`unified-inbox: calendar initial poll failed: ${String(err)}`),
    );

    this.interval = setInterval(async () => {
      if (this.shuttingDown || this.status.paused) return;
      await this.poll();
    }, this.cfg.calendar.pollIntervalMs);

    this.log.info(
      `unified-inbox: calendar monitor started (poll every ${this.cfg.calendar.pollIntervalMs / 1000}s)`,
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

  /** Fetch upcoming events for on-demand /calendar command */
  async fetchUpcomingEvents(): Promise<CalendarEvent[]> {
    const token = await this.auth.getAccessToken();
    const now = new Date();
    const endOfDay = new Date(now);
    endOfDay.setHours(23, 59, 59, 999);

    return fetchCalendarView(
      token,
      now.toISOString(),
      endOfDay.toISOString(),
    );
  }

  private async poll(): Promise<void> {
    try {
      const token = await this.auth.getAccessToken();

      const now = new Date();
      const lookAheadEnd = new Date(
        now.getTime() + this.cfg.calendar.lookAheadMinutes * 60 * 1000,
      );

      const events = await fetchCalendarView(
        token,
        now.toISOString(),
        lookAheadEnd.toISOString(),
      );

      for (const event of events) {
        if (event.isCancelled) continue;

        const eventStart = new Date(event.start.dateTime + "Z");
        const minutesUntilStart = (eventStart.getTime() - now.getTime()) / 60_000;

        // Check each reminder threshold
        for (const reminderMin of this.cfg.calendar.reminderMinutes) {
          if (minutesUntilStart <= reminderMin && minutesUntilStart > 0) {
            const sentReminders = this.notifiedEvents.get(event.id);
            if (!sentReminders?.has(reminderMin)) {
              await this.sendReminder(event, reminderMin);
              if (!this.notifiedEvents.has(event.id)) {
                this.notifiedEvents.set(event.id, new Set());
              }
              this.notifiedEvents.get(event.id)!.add(reminderMin);
            }
          }
        }
      }

      this.status.lastPollAt = Date.now();
      this.status.consecutiveFailures = 0;

      // Clean up old event tracking (older than lookAhead window)
      this.pruneNotifiedEvents();
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.status.lastErrorAt = Date.now();
      this.status.lastError = errorMsg;
      this.status.consecutiveFailures++;
      this.log.error(`unified-inbox: calendar poll error: ${errorMsg}`);

      if (this.status.consecutiveFailures >= 3) {
        this.status.paused = true;
        this.log.error(
          "unified-inbox: calendar monitor paused after 3 consecutive failures",
        );
        await sendTelegramMessage({
          botToken: this.cfg.telegramBotToken,
          chatId: this.cfg.telegramChatId,
          text: `[Unified Inbox] Calendar monitor paused after 3 failures. Last error: ${errorMsg}`,
        });
      }
    }
  }

  private async sendReminder(
    event: CalendarEvent,
    minutesBefore: number,
  ): Promise<void> {
    const text = formatCalendarReminder(event, minutesBefore);
    const joinUrl = event.onlineMeeting?.joinUrl || event.onlineMeetingUrl;
    await sendTelegramMessage({
      botToken: this.cfg.telegramBotToken,
      chatId: this.cfg.telegramChatId,
      text,
      parseMode: "MarkdownV2",
      replyMarkup: joinUrl
        ? { inline_keyboard: [[{ text: "Join Meeting", url: joinUrl }]] }
        : undefined,
    });

    // Push plain text to notification store for AI agent context
    pushNotification({
      source: "calendar",
      text: formatCalendarReminderPlain(event, minutesBefore),
      timestamp: Date.now(),
    });
  }

  private pruneNotifiedEvents(): void {
    // Keep only last 200 events
    if (this.notifiedEvents.size > 200) {
      const keys = Array.from(this.notifiedEvents.keys());
      for (const key of keys.slice(0, keys.length - 100)) {
        this.notifiedEvents.delete(key);
      }
    }
  }
}
