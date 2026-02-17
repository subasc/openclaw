// ============================================================================
// Format messages from each source for Telegram display
// ============================================================================

import type {
  EmailMessage,
  CalendarEvent,
  TeamsChatMessage,
  WhatsAppInboundMessage,
} from "./types.js";

/** Escape special characters for Telegram MarkdownV2 */
function esc(text: string): string {
  return text.replace(/([_*\[\]()~`>#+\-=|{}.!\\])/g, "\\$1");
}

/** Truncate text to a max length with ellipsis */
function truncate(text: string, max: number): string {
  if (text.length <= max) return text;
  return text.slice(0, max - 3) + "...";
}

/** Format a date string for display */
function formatTime(isoDate: string): string {
  const d = new Date(isoDate);
  return d.toLocaleTimeString("en-US", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  });
}

function formatDateTime(isoDate: string): string {
  const d = new Date(isoDate);
  return d.toLocaleString("en-US", {
    weekday: "short",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  });
}

// ============================================================================
// Email formatting
// ============================================================================

export function formatEmailNotification(email: EmailMessage): string {
  const from = email.from?.emailAddress;
  const fromName = from?.name || from?.address || "Unknown";
  const fromAddr = from?.address || "";
  const subject = email.subject || "(no subject)";
  const preview = truncate(email.bodyPreview || "", 500);
  const time = formatTime(email.receivedDateTime);
  const attachmentNote = email.hasAttachments ? "\nAttachments: yes" : "";

  return [
    `*\\[Email\\]* from *${esc(fromName)}*`,
    fromAddr ? `${esc(fromAddr)}` : null,
    `*Subject:* ${esc(subject)}`,
    `*Time:* ${esc(time)}${attachmentNote}`,
    `\\-\\-\\-`,
    esc(preview),
    ``,
    `_Reply to this message to respond_`,
  ]
    .filter((line) => line !== null)
    .join("\n");
}

export function formatEmailListItem(email: EmailMessage, index: number): string {
  const from = email.from?.emailAddress;
  const fromName = from?.name || from?.address || "Unknown";
  const subject = email.subject || "(no subject)";
  const time = formatTime(email.receivedDateTime);

  return `${index + 1}\\. *${esc(fromName)}* \\- ${esc(truncate(subject, 60))} \\(${esc(time)}\\)`;
}

// ============================================================================
// Calendar formatting
// ============================================================================

export function formatCalendarNotification(event: CalendarEvent): string {
  const startTime = formatDateTime(event.start.dateTime);
  const endTime = formatTime(event.end.dateTime);
  const location = event.location?.displayName;
  const organizer = event.organizer?.emailAddress?.name;
  const joinUrl =
    event.onlineMeeting?.joinUrl || event.onlineMeetingUrl;

  const lines = [
    `*\\[Calendar\\]* ${esc(event.subject)}`,
    event.isAllDay
      ? `*All day*`
      : `*Time:* ${esc(startTime)} \\- ${esc(endTime)}`,
  ];

  if (location) lines.push(`*Location:* ${esc(location)}`);
  if (organizer) lines.push(`*Organizer:* ${esc(organizer)}`);
  if (joinUrl) lines.push(`*Join:* ${esc(joinUrl)}`);
  if (event.bodyPreview) {
    lines.push(`\\-\\-\\-`);
    lines.push(esc(truncate(event.bodyPreview, 300)));
  }

  return lines.join("\n");
}

export function formatCalendarReminder(
  event: CalendarEvent,
  minutesBefore: number,
): string {
  const startTime = formatTime(event.start.dateTime);
  const joinUrl =
    event.onlineMeeting?.joinUrl || event.onlineMeetingUrl;

  const lines = [
    `*\\[Reminder\\]* ${esc(event.subject)} starts in *${minutesBefore} min*`,
    `*Time:* ${esc(startTime)}`,
  ];

  if (event.location?.displayName) {
    lines.push(`*Location:* ${esc(event.location.displayName)}`);
  }
  if (joinUrl) lines.push(`*Join:* ${esc(joinUrl)}`);

  return lines.join("\n");
}

export function formatCalendarListItem(
  event: CalendarEvent,
  index: number,
): string {
  const startTime = event.isAllDay
    ? "All day"
    : formatTime(event.start.dateTime);

  return `${index + 1}\\. *${esc(event.subject)}* \\- ${esc(startTime)}`;
}

// ============================================================================
// Teams chat formatting
// ============================================================================

export function formatTeamsChatNotification(
  msg: TeamsChatMessage,
  chatName: string,
): string {
  const sender =
    msg.from?.user?.displayName ||
    msg.from?.application?.displayName ||
    "Unknown";

  // Strip HTML tags from Teams message body
  const bodyText = stripHtml(msg.body?.content || "");
  const time = formatTime(msg.createdDateTime);

  return [
    `*\\[Teams\\]* ${esc(chatName)}`,
    `*${esc(sender)}* \\(${esc(time)}\\):`,
    esc(truncate(bodyText, 500)),
    ``,
    `_Reply to this message to respond in Teams_`,
  ].join("\n");
}

// ============================================================================
// WhatsApp formatting
// ============================================================================

export function formatWhatsAppNotification(msg: WhatsAppInboundMessage): string {
  const sender = msg.senderName || msg.from || "Unknown";
  const lines = [`*\\[WhatsApp\\]* *${esc(sender)}*`];

  if (msg.groupName) {
    lines[0] += ` in *${esc(msg.groupName)}*`;
  }

  lines.push(esc(truncate(msg.content, 500)));
  lines.push(``);
  lines.push(`_Reply to this message to respond on WhatsApp_`);

  return lines.join("\n");
}

// ============================================================================
// Plain text formatters (for AI agent context injection)
// ============================================================================

export function formatEmailPlain(email: EmailMessage): string {
  const from = email.from?.emailAddress;
  const fromName = from?.name || from?.address || "Unknown";
  const fromAddr = from?.address || "";
  const subject = email.subject || "(no subject)";
  const preview = truncate(email.bodyPreview || "", 500);
  const time = formatTime(email.receivedDateTime);

  return [
    `[Email] From: ${fromName}${fromAddr ? ` <${fromAddr}>` : ""}`,
    `Subject: ${subject}`,
    `Time: ${time}${email.hasAttachments ? " | Attachments: yes" : ""}`,
    preview,
  ].join("\n");
}

export function formatCalendarReminderPlain(
  event: CalendarEvent,
  minutesBefore: number,
): string {
  const startTime = formatTime(event.start.dateTime);
  const location = event.location?.displayName;
  const joinUrl = event.onlineMeeting?.joinUrl || event.onlineMeetingUrl;

  const lines = [
    `[Calendar Reminder] "${event.subject}" starts in ${minutesBefore} min`,
    `Time: ${startTime}`,
  ];
  if (location) lines.push(`Location: ${location}`);
  if (joinUrl) lines.push(`Join: ${joinUrl}`);
  return lines.join("\n");
}

export function formatTeamsChatPlain(
  msg: TeamsChatMessage,
  chatName: string,
): string {
  const sender =
    msg.from?.user?.displayName ||
    msg.from?.application?.displayName ||
    "Unknown";
  const bodyText = stripHtml(msg.body?.content || "");
  const time = formatTime(msg.createdDateTime);

  return [
    `[Teams Chat] ${chatName}`,
    `${sender} (${time}):`,
    truncate(bodyText, 500),
  ].join("\n");
}

export function formatWhatsAppPlain(msg: WhatsAppInboundMessage): string {
  const sender = msg.senderName || msg.from || "Unknown";
  const group = msg.groupName ? ` in ${msg.groupName}` : "";
  return `[WhatsApp] ${sender}${group}: ${truncate(msg.content, 500)}`;
}

// ============================================================================
// Helpers
// ============================================================================

/** Strip HTML tags from a string */
function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .trim();
}
