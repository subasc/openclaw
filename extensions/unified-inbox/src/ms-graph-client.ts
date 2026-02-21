// ============================================================================
// Microsoft Graph API client for email, calendar, and Teams chat
// Pattern follows extensions/msteams/src/graph.ts fetchGraphJson()
// ============================================================================

import type {
  EmailMessage,
  CalendarEvent,
  TeamsChat,
  TeamsChatMessage,
  TodoTaskList,
  TodoTask,
} from "./types.js";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

type GraphListResponse<T> = {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
};

async function fetchGraph<T>(params: {
  token: string;
  path: string;
  method?: string;
  body?: unknown;
  headers?: Record<string, string>;
}): Promise<T> {
  const url = params.path.startsWith("http") ? params.path : `${GRAPH_ROOT}${params.path}`;

  const res = await fetch(url, {
    method: params.method ?? "GET",
    headers: {
      Authorization: `Bearer ${params.token}`,
      "Content-Type": "application/json",
      ...params.headers,
    },
    body: params.body ? JSON.stringify(params.body) : undefined,
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph API ${res.status}: ${text.slice(0, 500)}`);
  }

  // Some endpoints return 204 No Content
  if (res.status === 204) return undefined as T;

  return (await res.json()) as T;
}

// ============================================================================
// Email endpoints
// ============================================================================

/** Fetch new/changed messages using delta query. Returns messages + deltaLink for next call. */
export async function fetchMailDelta(
  token: string,
  deltaLinkOrFolder: string,
  opts?: { filterUnread?: boolean; top?: number },
): Promise<{ messages: EmailMessage[]; deltaLink: string }> {
  const isFullDeltaLink = deltaLinkOrFolder.startsWith("http");

  let url: string;
  if (isFullDeltaLink) {
    url = deltaLinkOrFolder;
  } else {
    const folder = encodeURIComponent(deltaLinkOrFolder);
    const params = new URLSearchParams({
      $select:
        "id,conversationId,subject,bodyPreview,body,from,toRecipients,receivedDateTime,isRead,hasAttachments,importance",
      $orderby: "receivedDateTime desc",
      $top: String(opts?.top ?? 10),
    });
    if (opts?.filterUnread) {
      // The delta endpoint doesn't support $filter=isRead — use the regular messages endpoint.
      params.set("$filter", "isRead eq false");
      url = `/me/mailFolders/${folder}/messages?${params}`;
    } else {
      url = `/me/mailFolders/${folder}/messages/delta?${params}`;
    }
  }

  const allMessages: EmailMessage[] = [];
  let nextLink: string | undefined = url;
  let deltaLink = "";

  while (nextLink) {
    const res = await fetchGraph<GraphListResponse<EmailMessage>>({
      token,
      path: nextLink,
    });
    allMessages.push(...res.value);
    nextLink = res["@odata.nextLink"];
    if (res["@odata.deltaLink"]) {
      deltaLink = res["@odata.deltaLink"];
    }
  }

  return { messages: allMessages, deltaLink };
}

/** Reply to an email message */
export async function replyToEmail(token: string, messageId: string, body: string): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/messages/${messageId}/reply`,
    method: "POST",
    body: {
      comment: body,
    },
  });
}

/** Send a new email */
export async function sendMail(
  token: string,
  params: { to: string; subject: string; body: string },
): Promise<void> {
  await fetchGraph<void>({
    token,
    path: "/me/sendMail",
    method: "POST",
    body: {
      message: {
        subject: params.subject,
        body: { contentType: "Text", content: params.body },
        toRecipients: [{ emailAddress: { address: params.to } }],
      },
      saveToSentItems: true,
    },
  });
}

/** Mark an email as read */
export async function markAsRead(token: string, messageId: string): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/messages/${messageId}`,
    method: "PATCH",
    body: { isRead: true },
  });
}

// ============================================================================
// Calendar endpoints
// ============================================================================

/** Fetch calendar events within a time range */
export async function fetchCalendarView(
  token: string,
  startDateTime: string,
  endDateTime: string,
): Promise<CalendarEvent[]> {
  const params = new URLSearchParams({
    startDateTime,
    endDateTime,
    $select:
      "id,subject,start,end,location,organizer,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,bodyPreview,isAllDay,isCancelled",
    $orderby: "start/dateTime",
    $top: "50",
  });

  const res = await fetchGraph<GraphListResponse<CalendarEvent>>({
    token,
    path: `/me/calendarview?${params}`,
  });

  return res.value;
}

/** Create a new calendar event */
export async function createCalendarEvent(
  token: string,
  params: {
    subject: string;
    start: { dateTime: string; timeZone: string };
    end: { dateTime: string; timeZone: string };
    body?: string;
    location?: string;
    attendees?: Array<{ email: string; name?: string }>;
    isOnlineMeeting?: boolean;
  },
): Promise<CalendarEvent> {
  return fetchGraph<CalendarEvent>({
    token,
    path: "/me/events",
    method: "POST",
    body: {
      subject: params.subject,
      start: params.start,
      end: params.end,
      body: params.body ? { contentType: "Text", content: params.body } : undefined,
      location: params.location ? { displayName: params.location } : undefined,
      attendees: params.attendees?.map((a) => ({
        emailAddress: { address: a.email, name: a.name ?? a.email },
        type: "required",
      })),
      isOnlineMeeting: params.isOnlineMeeting ?? false,
    },
  });
}

/** Accept a calendar event invitation */
export async function acceptCalendarEvent(
  token: string,
  eventId: string,
  comment?: string,
): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/events/${eventId}/accept`,
    method: "POST",
    body: { comment: comment ?? "", sendResponse: true },
  });
}

/** Decline a calendar event invitation */
export async function declineCalendarEvent(
  token: string,
  eventId: string,
  comment?: string,
): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/events/${eventId}/decline`,
    method: "POST",
    body: { comment: comment ?? "", sendResponse: true },
  });
}

/** Tentatively accept a calendar event invitation */
export async function tentativelyAcceptCalendarEvent(
  token: string,
  eventId: string,
  comment?: string,
): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/events/${eventId}/tentativelyAccept`,
    method: "POST",
    body: { comment: comment ?? "", sendResponse: true },
  });
}

/** Delete a calendar event */
export async function deleteCalendarEvent(token: string, eventId: string): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/events/${eventId}`,
    method: "DELETE",
  });
}

// ============================================================================
// ToDo / Tasks endpoints
// ============================================================================

/** List all ToDo task lists */
export async function listTodoLists(token: string): Promise<TodoTaskList[]> {
  const res = await fetchGraph<GraphListResponse<TodoTaskList>>({
    token,
    path: "/me/todo/lists?$top=50",
  });
  return res.value;
}

/** List tasks in a ToDo list */
export async function listTodoTasks(
  token: string,
  listId: string,
  opts?: { status?: "notStarted" | "inProgress" | "completed" | "all"; top?: number },
): Promise<TodoTask[]> {
  const params = new URLSearchParams({
    $top: String(opts?.top ?? 25),
    $orderby: "createdDateTime desc",
  });

  if (opts?.status && opts.status !== "all") {
    params.set("$filter", `status eq '${opts.status}'`);
  }

  const res = await fetchGraph<GraphListResponse<TodoTask>>({
    token,
    path: `/me/todo/lists/${listId}/tasks?${params}`,
  });
  return res.value;
}

/** Create a new ToDo task */
export async function createTodoTask(
  token: string,
  listId: string,
  params: {
    title: string;
    body?: string;
    importance?: "low" | "normal" | "high";
    dueDateTime?: { dateTime: string; timeZone: string };
    reminderDateTime?: { dateTime: string; timeZone: string };
  },
): Promise<TodoTask> {
  return fetchGraph<TodoTask>({
    token,
    path: `/me/todo/lists/${listId}/tasks`,
    method: "POST",
    body: {
      title: params.title,
      body: params.body ? { content: params.body, contentType: "text" } : undefined,
      importance: params.importance ?? "normal",
      dueDateTime: params.dueDateTime,
      reminderDateTime: params.reminderDateTime,
      isReminderOn: !!params.reminderDateTime,
    },
  });
}

/** Update a ToDo task */
export async function updateTodoTask(
  token: string,
  listId: string,
  taskId: string,
  updates: Partial<{
    title: string;
    status: "notStarted" | "inProgress" | "completed";
    importance: "low" | "normal" | "high";
    dueDateTime: { dateTime: string; timeZone: string } | null;
  }>,
): Promise<TodoTask> {
  return fetchGraph<TodoTask>({
    token,
    path: `/me/todo/lists/${listId}/tasks/${taskId}`,
    method: "PATCH",
    body: updates,
  });
}

/** Delete a ToDo task */
export async function deleteTodoTask(token: string, listId: string, taskId: string): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/todo/lists/${listId}/tasks/${taskId}`,
    method: "DELETE",
  });
}

// ============================================================================
// Teams Chat endpoints
// ============================================================================

/** List recent chats with last message preview */
export async function listChats(token: string): Promise<TeamsChat[]> {
  const params = new URLSearchParams({
    $expand: "lastMessagePreview",
    $orderby: "lastMessagePreview/createdDateTime desc",
    $top: "20",
  });

  const res = await fetchGraph<GraphListResponse<TeamsChat>>({
    token,
    path: `/me/chats?${params}`,
  });

  return res.value;
}

/** List messages in a specific chat since a given timestamp */
export async function listChatMessages(
  token: string,
  chatId: string,
  opts?: { since?: string; top?: number },
): Promise<TeamsChatMessage[]> {
  const params = new URLSearchParams({
    $top: String(opts?.top ?? 20),
    $orderby: "createdDateTime desc",
  });

  const res = await fetchGraph<GraphListResponse<TeamsChatMessage>>({
    token,
    path: `/me/chats/${chatId}/messages?${params}`,
  });

  // Graph API doesn't support $filter on createdDateTime for chat messages,
  // so filter client-side instead.
  if (opts?.since) {
    return res.value.filter((msg) => msg.createdDateTime > opts.since!);
  }

  return res.value;
}

/** Send a message in a Teams chat */
export async function sendChatMessage(token: string, chatId: string, body: string): Promise<void> {
  await fetchGraph<void>({
    token,
    path: `/me/chats/${chatId}/messages`,
    method: "POST",
    body: {
      body: { contentType: "text", content: body },
    },
  });
}

// ============================================================================
// Outlook REST API (uses outlook.office.com token which has Mail.Send scope)
// ============================================================================

const OUTLOOK_REST_ROOT = "https://outlook.office.com/api/v2.0";

/** Send a new email via the Outlook REST API (requires outlook.office.com token with Mail.Send) */
export async function sendMailViaOutlookRest(
  token: string,
  params: { to: string; subject: string; body: string },
): Promise<void> {
  const res = await fetch(`${OUTLOOK_REST_ROOT}/me/sendmail`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      Message: {
        Subject: params.subject,
        Body: { ContentType: "Text", Content: params.body },
        ToRecipients: [{ EmailAddress: { Address: params.to } }],
      },
      SaveToSentItems: true,
    }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Outlook REST API ${res.status}: ${text.slice(0, 500)}`);
  }
}

/** Reply to an email via the Outlook REST API */
export async function replyToEmailViaOutlookRest(
  token: string,
  messageId: string,
  body: string,
): Promise<void> {
  const res = await fetch(`${OUTLOOK_REST_ROOT}/me/messages/${messageId}/reply`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      Comment: body,
    }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Outlook REST API ${res.status}: ${text.slice(0, 500)}`);
  }
}

/** Reply-all to an email via the Outlook REST API */
export async function replyAllToEmailViaOutlookRest(
  token: string,
  messageId: string,
  body: string,
): Promise<void> {
  const res = await fetch(`${OUTLOOK_REST_ROOT}/me/messages/${messageId}/replyall`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      Comment: body,
    }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Outlook REST API ${res.status}: ${text.slice(0, 500)}`);
  }
}

// ============================================================================
// User profile
// ============================================================================

/** Get current user's profile */
export async function getMe(
  token: string,
): Promise<{ displayName: string; mail: string; id: string }> {
  return fetchGraph({
    token,
    path: "/me?$select=displayName,mail,id",
  });
}
