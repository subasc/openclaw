// ============================================================================
// Shared types for the Unified Inbox extension
// ============================================================================

/** Microsoft Graph email message (subset of fields we use) */
export type EmailMessage = {
  id: string;
  conversationId?: string;
  subject: string;
  bodyPreview: string;
  body?: { contentType: string; content: string };
  from: { emailAddress: { name: string; address: string } };
  toRecipients?: Array<{ emailAddress: { name: string; address: string } }>;
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance?: "low" | "normal" | "high";
};

/** Microsoft Graph calendar event */
export type CalendarEvent = {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName?: string };
  organizer?: { emailAddress: { name: string; address: string } };
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  onlineMeeting?: { joinUrl?: string };
  bodyPreview?: string;
  isAllDay: boolean;
  isCancelled: boolean;
};

/** Microsoft Graph Teams chat message */
export type TeamsChatMessage = {
  id: string;
  createdDateTime: string;
  body: { contentType: string; content: string };
  from?: {
    user?: { id: string; displayName: string };
    application?: { id: string; displayName: string };
  };
  chatId: string;
  messageType: string;
};

/** Microsoft Graph Teams chat */
export type TeamsChat = {
  id: string;
  topic?: string;
  chatType: "oneOnOne" | "group" | "meeting";
  lastMessagePreview?: {
    id: string;
    createdDateTime: string;
    body: { content: string };
    from?: {
      user?: { displayName: string };
      application?: { displayName: string };
    };
  };
};

/** Microsoft Graph ToDo task list */
export type TodoTaskList = {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
  wellknownListName?: "none" | "defaultList" | "flaggedEmails" | "unknownFutureValue";
};

/** Microsoft Graph ToDo task */
export type TodoTask = {
  id: string;
  title: string;
  body?: { content: string; contentType: string };
  status: "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred";
  importance: "low" | "normal" | "high";
  isReminderOn: boolean;
  createdDateTime: string;
  lastModifiedDateTime: string;
  completedDateTime?: { dateTime: string; timeZone: string };
  dueDateTime?: { dateTime: string; timeZone: string };
  reminderDateTime?: { dateTime: string; timeZone: string };
  categories?: string[];
};

/** Inbound WhatsApp message (from the message_received hook) */
export type WhatsAppInboundMessage = {
  from: string;
  content: string;
  senderName?: string;
  groupName?: string;
  groupJid?: string;
  messageId?: string;
};

/** Source types for reply routing */
export type ReplySource = "email" | "teams-chat" | "whatsapp";

/** Context stored per forwarded message for routing replies back */
export type ReplySourceContext =
  | {
      type: "email";
      messageId: string;
      conversationId?: string;
      fromAddress: string;
      subject: string;
      fromName?: string;
      bodyPreview?: string;
      toRecipients?: Array<{ name: string; address: string }>;
    }
  | {
      type: "teams-chat";
      chatId: string;
      messageId?: string;
    }
  | {
      type: "whatsapp";
      jid: string;
      messageId?: string;
      participant?: string;
    };

/** A single reply mapping entry */
export type ReplyMapping = {
  telegramMessageId: number;
  source: ReplySource;
  sourceContext: ReplySourceContext;
  createdAt: number;
};

/** Common auth provider interface for both device-code and browser-based auth */
export interface IMsAuthProvider {
  isAuthenticated(): boolean;
  getAccessToken(): Promise<string>;
  loadPersistedTokens(): Promise<boolean>;
  startAutoRefresh(onError: (error: string) => Promise<void>): void;
  stopAutoRefresh(): void;
  /** Register callback invoked when token recovers after failure */
  setOnTokenRecovered?(cb: () => void): void;
  /** Force token refresh/recovery (for remote management commands) */
  forceRefresh?(): Promise<boolean>;
}

/** Token data persisted to disk */
export type TokenData = {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  account?: string;
  scopes: string[];
};

/** Context for Teams inline button callbacks */
export type TeamsButtonContext = {
  type: "teams";
  chatId: string;
  chatName: string;
  senderName: string;
  messageId?: string;
  telegramMessageId: number;
  registeredAt: number;
};

/** Context for WhatsApp inline button callbacks */
export type WhatsAppButtonContext = {
  type: "whatsapp";
  jid: string;
  senderName: string;
  content: string;
  groupName?: string;
  groupJid?: string;
  messageId?: string;
  telegramMessageId: number;
  registeredAt: number;
};

/** Union of all button contexts for the shared ShortIdRegistry */
export type ButtonContext =
  | ({ type: "email" } & {
      messageId: string;
      fromName: string;
      fromAddress: string;
      subject: string;
      bodyPreview: string;
      toRecipients: Array<{ name: string; address: string }>;
      telegramMessageId: number;
      registeredAt: number;
    })
  | TeamsButtonContext
  | WhatsAppButtonContext;

/** Monitor status for /inbox-status command */
export type MonitorStatus = {
  running: boolean;
  lastPollAt: number | null;
  lastErrorAt: number | null;
  lastError: string | null;
  consecutiveFailures: number;
  paused: boolean;
};
