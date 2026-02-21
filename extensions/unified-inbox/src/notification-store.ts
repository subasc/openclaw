// ============================================================================
// In-memory ring buffer of recent notifications for AI agent context injection.
// Monitors push plain-text summaries here; the before_agent_start hook reads
// them and prepends them as context so the AI can reference emails/events/chats.
// ============================================================================

import { persistToLongTermMemory } from "./memory-bridge.js";

export type NotificationEntry = {
  source: "email" | "calendar" | "teams" | "whatsapp";
  text: string;
  timestamp: number;
};

const MAX_ENTRIES = 50;
const MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24 hours

let entries: NotificationEntry[] = [];

export function pushNotification(entry: NotificationEntry): void {
  entries.push(entry);
  if (entries.length > MAX_ENTRIES) {
    entries = entries.slice(-MAX_ENTRIES);
  }
  // Fire-and-forget: persist to vector memory for long-term recall
  persistToLongTermMemory(entry.text, entry.source).catch(() => {});
}

export function getRecentNotifications(): NotificationEntry[] {
  const cutoff = Date.now() - MAX_AGE_MS;
  entries = entries.filter((e) => e.timestamp >= cutoff);
  return entries;
}

export function formatNotificationsAsContext(): string | undefined {
  const recent = getRecentNotifications();
  if (recent.length === 0) return undefined;

  const lines = recent.map((e) => e.text);
  return [
    "=== Recent Notifications (from Unified Inbox) ===",
    "The following notifications were forwarded to the user's Telegram chat.",
    "Use these to answer questions about recent emails, calendar events, and Teams messages.",
    "",
    ...lines,
    "",
    "=== End Notifications ===",
  ].join("\n");
}
