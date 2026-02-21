// ============================================================================
// Agent tools: email, calendar, ToDo & Teams actions the AI can invoke
// ============================================================================

import { Type } from "@sinclair/typebox";
import type { IMsAuthProvider } from "./types.js";
import {
  sendMailViaOutlookRest,
  replyToEmailViaOutlookRest,
  sendChatMessage,
  listChats,
  fetchMailDelta,
  fetchCalendarView,
  createCalendarEvent,
  acceptCalendarEvent,
  declineCalendarEvent,
  tentativelyAcceptCalendarEvent,
  deleteCalendarEvent,
  listTodoLists,
  listTodoTasks,
  createTodoTask,
  updateTodoTask,
  deleteTodoTask,
} from "./ms-graph-client.js";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

// Shared auth providers — set by service.ts after startup
let mailAuth: IMsAuthProvider | null = null;
let mailReadAuth: IMsAuthProvider | null = null;
let calendarAuth: IMsAuthProvider | null = null;
let chatAuth: IMsAuthProvider | null = null;
let toolLogger: Logger | null = null;

export function setToolAuthProviders(deps: {
  mailAuth: IMsAuthProvider;
  mailReadAuth?: IMsAuthProvider;
  calendarAuth?: IMsAuthProvider;
  chatAuth: IMsAuthProvider;
  log: Logger;
}): void {
  mailAuth = deps.mailAuth;
  mailReadAuth = deps.mailReadAuth ?? deps.mailAuth;
  calendarAuth = deps.calendarAuth ?? deps.mailAuth;
  chatAuth = deps.chatAuth;
  toolLogger = deps.log;
}

function json(payload: unknown) {
  return {
    content: [{ type: "text" as const, text: JSON.stringify(payload, null, 2) }],
    details: payload,
  };
}

// ============================================================================
// list_emails tool
// ============================================================================

export const listEmailsTool = {
  name: "list_emails",
  label: "List Recent Emails",
  description:
    "Fetch recent emails from the user's Microsoft 365 inbox. Use this when the user asks about their emails, new messages, or wants to check their inbox.",
  parameters: Type.Object({
    count: Type.Optional(
      Type.Number({ description: "Number of emails to fetch (default 10, max 20)" }),
    ),
    unreadOnly: Type.Optional(
      Type.Boolean({ description: "Only return unread emails (default true)" }),
    ),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!mailReadAuth)
      throw new Error("Email auth not available — unified-inbox service may not be running");
    const count = Math.min(Math.max(Number(params.count) || 10, 1), 20);
    const unreadOnly = params.unreadOnly !== false;

    const token = await mailReadAuth.getAccessToken();
    const result = await fetchMailDelta(token, "Inbox", {
      filterUnread: unreadOnly,
      top: count,
    });

    const emails = result.messages.map((msg) => ({
      id: msg.id,
      from: msg.from?.emailAddress?.name
        ? `${msg.from.emailAddress.name} <${msg.from.emailAddress.address}>`
        : (msg.from?.emailAddress?.address ?? "Unknown"),
      subject: msg.subject ?? "(no subject)",
      preview: (msg.bodyPreview ?? "").slice(0, 200),
      time: msg.receivedDateTime,
      isRead: msg.isRead,
      hasAttachments: msg.hasAttachments,
      importance: msg.importance,
      to: (msg.toRecipients ?? []).map((r) =>
        r.emailAddress?.name
          ? `${r.emailAddress.name} <${r.emailAddress.address}>`
          : (r.emailAddress?.address ?? ""),
      ),
    }));

    toolLogger?.info(`unified-inbox: tool listed ${emails.length} emails`);
    return json({ count: emails.length, emails });
  },
};

// ============================================================================
// send_email tool
// ============================================================================

export const sendEmailTool = {
  name: "send_email",
  label: "Send Email",
  description:
    "Send a new email via Microsoft 365. Use this when the user asks you to send, compose, or write an email.",
  parameters: Type.Object({
    to: Type.String({ description: "Recipient email address" }),
    subject: Type.String({ description: "Email subject line" }),
    body: Type.String({ description: "Email body text" }),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!mailAuth)
      throw new Error("Email auth not available — unified-inbox service may not be running");
    const to = String(params.to ?? "").trim();
    const subject = String(params.subject ?? "").trim();
    const body = String(params.body ?? "").trim();
    if (!to) throw new Error("'to' address is required");
    if (!subject) throw new Error("'subject' is required");
    if (!body) throw new Error("'body' is required");

    const token = await mailAuth.getAccessToken();
    await sendMailViaOutlookRest(token, { to, subject, body });
    toolLogger?.info(`unified-inbox: tool sent email to ${to} — "${subject}"`);
    return json({ success: true, to, subject });
  },
};

// ============================================================================
// reply_email tool
// ============================================================================

export const replyEmailTool = {
  name: "reply_email",
  label: "Reply to Email",
  description:
    "Reply to an existing email by its Graph message ID. The message ID is available from recent email notifications.",
  parameters: Type.Object({
    messageId: Type.String({
      description: "The Microsoft Graph message ID of the email to reply to",
    }),
    body: Type.String({ description: "Reply body text" }),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!mailAuth) throw new Error("Email auth not available");
    const messageId = String(params.messageId ?? "").trim();
    const body = String(params.body ?? "").trim();
    if (!messageId) throw new Error("'messageId' is required");
    if (!body) throw new Error("'body' is required");

    const token = await mailAuth.getAccessToken();
    await replyToEmailViaOutlookRest(token, messageId, body);
    toolLogger?.info(`unified-inbox: tool replied to email ${messageId}`);
    return json({ success: true, messageId });
  },
};

// ============================================================================
// list_teams_chats tool
// ============================================================================

export const listTeamsChatsTool = {
  name: "list_teams_chats",
  label: "List Teams Chats",
  description:
    "List recent Microsoft Teams chats. Returns chat IDs, names, and last message previews. Use this to find the chat ID before sending a Teams message.",
  parameters: Type.Object({}),
  async execute(_id: string, _params: Record<string, unknown>) {
    if (!chatAuth) throw new Error("Teams auth not available");
    const token = await chatAuth.getAccessToken();
    const chats = await listChats(token);

    const summary = chats.map((c) => ({
      chatId: c.id,
      topic: c.topic || null,
      chatType: c.chatType,
      lastMessage: c.lastMessagePreview
        ? {
            from: c.lastMessagePreview.from?.user?.displayName ?? "Unknown",
            preview: (c.lastMessagePreview.body?.content ?? "")
              .replace(/<[^>]+>/g, "")
              .slice(0, 120),
            time: c.lastMessagePreview.createdDateTime,
          }
        : null,
    }));

    toolLogger?.info(`unified-inbox: tool listed ${summary.length} Teams chats`);
    return json({ chats: summary });
  },
};

// ============================================================================
// send_teams_message tool
// ============================================================================

export const sendTeamsMessageTool = {
  name: "send_teams_message",
  label: "Send Teams Message",
  description:
    "Send a message in a Microsoft Teams chat. You need the chat ID — use list_teams_chats first to find it. Use this when the user asks you to message someone on Teams.",
  parameters: Type.Object({
    chatId: Type.String({ description: "The Teams chat ID to send the message to" }),
    message: Type.String({ description: "Message text to send" }),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!chatAuth) throw new Error("Teams auth not available");
    const chatId = String(params.chatId ?? "").trim();
    const message = String(params.message ?? "").trim();
    if (!chatId) throw new Error("'chatId' is required");
    if (!message) throw new Error("'message' is required");

    const token = await chatAuth.getAccessToken();
    await sendChatMessage(token, chatId, message);
    toolLogger?.info(`unified-inbox: tool sent Teams message to chat ${chatId}`);
    return json({ success: true, chatId });
  },
};

// ============================================================================
// list_calendar tool
// ============================================================================

export const listCalendarTool = {
  name: "list_calendar",
  label: "List Calendar Events",
  description:
    "Fetch upcoming calendar events from Microsoft 365. Use this when the user asks about their schedule, meetings, or calendar. Returns events for today by default, or a custom date range.",
  parameters: Type.Object({
    startDate: Type.Optional(
      Type.String({
        description:
          "Start date/time in ISO 8601 format (e.g. '2026-02-21T00:00:00'). Defaults to now.",
      }),
    ),
    endDate: Type.Optional(
      Type.String({
        description:
          "End date/time in ISO 8601 format (e.g. '2026-02-21T23:59:59'). Defaults to end of today.",
      }),
    ),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth)
      throw new Error("Calendar auth not available — unified-inbox service may not be running");

    const now = new Date();
    const startDate = params.startDate ? String(params.startDate) : now.toISOString();
    const endOfDay = new Date(now);
    endOfDay.setHours(23, 59, 59, 999);
    const endDate = params.endDate ? String(params.endDate) : endOfDay.toISOString();

    const token = await calendarAuth.getAccessToken();
    const events = await fetchCalendarView(token, startDate, endDate);

    const summary = events
      .filter((e) => !e.isCancelled)
      .map((e) => ({
        id: e.id,
        subject: e.subject,
        start: e.start.dateTime,
        end: e.end.dateTime,
        location: e.location?.displayName || null,
        organizer: e.organizer?.emailAddress?.name || null,
        isOnlineMeeting: e.isOnlineMeeting,
        joinUrl: e.onlineMeeting?.joinUrl || e.onlineMeetingUrl || null,
        isAllDay: e.isAllDay,
        preview: e.bodyPreview?.slice(0, 200) || null,
      }));

    toolLogger?.info(`unified-inbox: tool listed ${summary.length} calendar events`);
    return json({ count: summary.length, events: summary });
  },
};

// ============================================================================
// create_event tool
// ============================================================================

export const createEventTool = {
  name: "create_event",
  label: "Create Calendar Event",
  description:
    "Create a new calendar event in Microsoft 365. Use this when the user asks you to schedule a meeting, add an event, or block time on their calendar.",
  parameters: Type.Object({
    subject: Type.String({ description: "Event title/subject" }),
    startDateTime: Type.String({
      description: "Start date and time in ISO 8601 format (e.g. '2026-02-21T14:00:00')",
    }),
    endDateTime: Type.String({
      description: "End date and time in ISO 8601 format (e.g. '2026-02-21T15:00:00')",
    }),
    timeZone: Type.Optional(
      Type.String({
        description: "Time zone (e.g. 'Asia/Kuala_Lumpur'). Defaults to Asia/Kuala_Lumpur.",
      }),
    ),
    body: Type.Optional(Type.String({ description: "Event description/body text" })),
    location: Type.Optional(Type.String({ description: "Event location" })),
    attendees: Type.Optional(
      Type.Array(Type.String({ description: "Attendee email address" }), {
        description: "List of attendee email addresses to invite",
      }),
    ),
    isOnlineMeeting: Type.Optional(
      Type.Boolean({ description: "Create as Teams online meeting (default false)" }),
    ),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("Calendar auth not available");
    const subject = String(params.subject ?? "").trim();
    if (!subject) throw new Error("'subject' is required");

    const tz = String(params.timeZone ?? "Asia/Kuala_Lumpur");
    const startDateTime = String(params.startDateTime ?? "").trim();
    const endDateTime = String(params.endDateTime ?? "").trim();
    if (!startDateTime) throw new Error("'startDateTime' is required");
    if (!endDateTime) throw new Error("'endDateTime' is required");

    const attendees = Array.isArray(params.attendees)
      ? (params.attendees as string[]).map((email) => ({ email: String(email).trim() }))
      : undefined;

    const token = await calendarAuth.getAccessToken();
    const event = await createCalendarEvent(token, {
      subject,
      start: { dateTime: startDateTime, timeZone: tz },
      end: { dateTime: endDateTime, timeZone: tz },
      body: params.body ? String(params.body) : undefined,
      location: params.location ? String(params.location) : undefined,
      attendees,
      isOnlineMeeting: Boolean(params.isOnlineMeeting),
    });

    toolLogger?.info(`unified-inbox: tool created event "${subject}" (${event.id})`);
    return json({
      success: true,
      eventId: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
    });
  },
};

// ============================================================================
// respond_event tool
// ============================================================================

export const respondEventTool = {
  name: "respond_event",
  label: "Respond to Calendar Invitation",
  description:
    "Accept, decline, or tentatively accept a calendar event invitation. Use this when the user wants to RSVP to a meeting. Use list_calendar first to find the event ID.",
  parameters: Type.Object({
    eventId: Type.String({ description: "The calendar event ID to respond to" }),
    response: Type.Union(
      [Type.Literal("accept"), Type.Literal("decline"), Type.Literal("tentative")],
      { description: "Response: 'accept', 'decline', or 'tentative'" },
    ),
    comment: Type.Optional(Type.String({ description: "Optional message to the organizer" })),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("Calendar auth not available");
    const eventId = String(params.eventId ?? "").trim();
    const response = String(params.response ?? "").trim();
    if (!eventId) throw new Error("'eventId' is required");
    if (!["accept", "decline", "tentative"].includes(response))
      throw new Error("'response' must be 'accept', 'decline', or 'tentative'");

    const comment = params.comment ? String(params.comment) : undefined;
    const token = await calendarAuth.getAccessToken();

    if (response === "accept") {
      await acceptCalendarEvent(token, eventId, comment);
    } else if (response === "decline") {
      await declineCalendarEvent(token, eventId, comment);
    } else {
      await tentativelyAcceptCalendarEvent(token, eventId, comment);
    }

    toolLogger?.info(`unified-inbox: tool ${response}ed event ${eventId}`);
    return json({ success: true, eventId, response });
  },
};

// ============================================================================
// delete_event tool
// ============================================================================

export const deleteEventTool = {
  name: "delete_event",
  label: "Delete Calendar Event",
  description:
    "Delete a calendar event. Use this when the user wants to cancel or remove a meeting from their calendar.",
  parameters: Type.Object({
    eventId: Type.String({ description: "The calendar event ID to delete" }),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("Calendar auth not available");
    const eventId = String(params.eventId ?? "").trim();
    if (!eventId) throw new Error("'eventId' is required");

    const token = await calendarAuth.getAccessToken();
    await deleteCalendarEvent(token, eventId);

    toolLogger?.info(`unified-inbox: tool deleted event ${eventId}`);
    return json({ success: true, eventId });
  },
};

// ============================================================================
// list_tasks tool
// ============================================================================

export const listTasksTool = {
  name: "list_tasks",
  label: "List ToDo Tasks",
  description:
    "List tasks from Microsoft To Do. Use this when the user asks about their tasks, to-do items, or what they need to do. Returns tasks from the default list unless a specific list is requested.",
  parameters: Type.Object({
    listName: Type.Optional(
      Type.String({
        description:
          "Name of the task list to fetch from. Omit to auto-detect the default 'Tasks' list.",
      }),
    ),
    status: Type.Optional(
      Type.Union(
        [
          Type.Literal("notStarted"),
          Type.Literal("inProgress"),
          Type.Literal("completed"),
          Type.Literal("all"),
        ],
        { description: "Filter by status (default: 'notStarted')" },
      ),
    ),
    count: Type.Optional(
      Type.Number({ description: "Number of tasks to fetch (default 25, max 50)" }),
    ),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth)
      throw new Error("ToDo auth not available — unified-inbox service may not be running");

    const token = await calendarAuth.getAccessToken();

    // Find the right list
    const lists = await listTodoLists(token);
    let targetList = lists.find(
      (l) => l.wellknownListName === "defaultList" || l.displayName === "Tasks",
    );

    if (params.listName) {
      const name = String(params.listName).toLowerCase();
      targetList = lists.find((l) => l.displayName.toLowerCase().includes(name)) ?? targetList;
    }

    if (!targetList) {
      return json({
        count: 0,
        tasks: [],
        availableLists: lists.map((l) => l.displayName),
        error: "No matching task list found",
      });
    }

    const status = (params.status as string) ?? "notStarted";
    const count = Math.min(Math.max(Number(params.count) || 25, 1), 50);

    const tasks = await listTodoTasks(token, targetList.id, {
      status: status as "notStarted" | "inProgress" | "completed" | "all",
      top: count,
    });

    const summary = tasks.map((t) => ({
      id: t.id,
      listId: targetList!.id,
      title: t.title,
      status: t.status,
      importance: t.importance,
      dueDate: t.dueDateTime?.dateTime || null,
      createdAt: t.createdDateTime,
      body: t.body?.content?.slice(0, 200) || null,
    }));

    toolLogger?.info(
      `unified-inbox: tool listed ${summary.length} tasks from "${targetList.displayName}"`,
    );
    return json({
      listName: targetList.displayName,
      listId: targetList.id,
      count: summary.length,
      tasks: summary,
    });
  },
};

// ============================================================================
// create_task tool
// ============================================================================

export const createTaskTool = {
  name: "create_task",
  label: "Create ToDo Task",
  description:
    "Create a new task in Microsoft To Do. Use this when the user asks you to add a task, reminder, or to-do item.",
  parameters: Type.Object({
    title: Type.String({ description: "Task title" }),
    body: Type.Optional(Type.String({ description: "Task description/notes" })),
    importance: Type.Optional(
      Type.Union([Type.Literal("low"), Type.Literal("normal"), Type.Literal("high")], {
        description: "Task importance (default: 'normal')",
      }),
    ),
    dueDate: Type.Optional(
      Type.String({
        description: "Due date in ISO 8601 format (e.g. '2026-02-22T17:00:00')",
      }),
    ),
    reminderDate: Type.Optional(
      Type.String({
        description: "Reminder date/time in ISO 8601 format",
      }),
    ),
    listName: Type.Optional(Type.String({ description: "Task list name (default: 'Tasks')" })),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("ToDo auth not available");
    const title = String(params.title ?? "").trim();
    if (!title) throw new Error("'title' is required");

    const token = await calendarAuth.getAccessToken();
    const tz = "Asia/Kuala_Lumpur";

    // Find the right list
    const lists = await listTodoLists(token);
    let targetList = lists.find(
      (l) => l.wellknownListName === "defaultList" || l.displayName === "Tasks",
    );

    if (params.listName) {
      const name = String(params.listName).toLowerCase();
      targetList = lists.find((l) => l.displayName.toLowerCase().includes(name)) ?? targetList;
    }

    if (!targetList) throw new Error("No task list found");

    const task = await createTodoTask(token, targetList.id, {
      title,
      body: params.body ? String(params.body) : undefined,
      importance: (params.importance as "low" | "normal" | "high") ?? "normal",
      dueDateTime: params.dueDate ? { dateTime: String(params.dueDate), timeZone: tz } : undefined,
      reminderDateTime: params.reminderDate
        ? { dateTime: String(params.reminderDate), timeZone: tz }
        : undefined,
    });

    toolLogger?.info(`unified-inbox: tool created task "${title}" in "${targetList.displayName}"`);
    return json({
      success: true,
      taskId: task.id,
      listId: targetList.id,
      listName: targetList.displayName,
      title: task.title,
      status: task.status,
    });
  },
};

// ============================================================================
// complete_task tool
// ============================================================================

export const completeTaskTool = {
  name: "complete_task",
  label: "Complete ToDo Task",
  description:
    "Mark a Microsoft To Do task as completed. Use list_tasks first to find the task and list IDs.",
  parameters: Type.Object({
    listId: Type.String({ description: "The task list ID" }),
    taskId: Type.String({ description: "The task ID to mark as completed" }),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("ToDo auth not available");
    const listId = String(params.listId ?? "").trim();
    const taskId = String(params.taskId ?? "").trim();
    if (!listId) throw new Error("'listId' is required");
    if (!taskId) throw new Error("'taskId' is required");

    const token = await calendarAuth.getAccessToken();
    await updateTodoTask(token, listId, taskId, { status: "completed" });

    toolLogger?.info(`unified-inbox: tool completed task ${taskId}`);
    return json({ success: true, taskId, status: "completed" });
  },
};

// ============================================================================
// update_task tool
// ============================================================================

export const updateTaskTool = {
  name: "update_task",
  label: "Update ToDo Task",
  description:
    "Update a Microsoft To Do task (change title, importance, due date, or status). Use list_tasks first to find the task and list IDs.",
  parameters: Type.Object({
    listId: Type.String({ description: "The task list ID" }),
    taskId: Type.String({ description: "The task ID to update" }),
    title: Type.Optional(Type.String({ description: "New task title" })),
    importance: Type.Optional(
      Type.Union([Type.Literal("low"), Type.Literal("normal"), Type.Literal("high")], {
        description: "New importance level",
      }),
    ),
    status: Type.Optional(
      Type.Union(
        [Type.Literal("notStarted"), Type.Literal("inProgress"), Type.Literal("completed")],
        { description: "New status" },
      ),
    ),
    dueDate: Type.Optional(
      Type.String({ description: "New due date in ISO 8601 format, or 'none' to remove" }),
    ),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("ToDo auth not available");
    const listId = String(params.listId ?? "").trim();
    const taskId = String(params.taskId ?? "").trim();
    if (!listId) throw new Error("'listId' is required");
    if (!taskId) throw new Error("'taskId' is required");

    const updates: Record<string, unknown> = {};
    if (params.title) updates.title = String(params.title);
    if (params.importance) updates.importance = String(params.importance);
    if (params.status) updates.status = String(params.status);
    if (params.dueDate) {
      updates.dueDateTime =
        params.dueDate === "none"
          ? null
          : { dateTime: String(params.dueDate), timeZone: "Asia/Kuala_Lumpur" };
    }

    if (Object.keys(updates).length === 0) throw new Error("No updates provided");

    const token = await calendarAuth.getAccessToken();
    const task = await updateTodoTask(token, listId, taskId, updates as any);

    toolLogger?.info(`unified-inbox: tool updated task ${taskId}`);
    return json({ success: true, taskId: task.id, title: task.title, status: task.status });
  },
};

// ============================================================================
// delete_task tool
// ============================================================================

export const deleteTaskTool = {
  name: "delete_task",
  label: "Delete ToDo Task",
  description:
    "Delete a Microsoft To Do task permanently. Use list_tasks first to find the task and list IDs.",
  parameters: Type.Object({
    listId: Type.String({ description: "The task list ID" }),
    taskId: Type.String({ description: "The task ID to delete" }),
  }),
  async execute(_id: string, params: Record<string, unknown>) {
    if (!calendarAuth) throw new Error("ToDo auth not available");
    const listId = String(params.listId ?? "").trim();
    const taskId = String(params.taskId ?? "").trim();
    if (!listId) throw new Error("'listId' is required");
    if (!taskId) throw new Error("'taskId' is required");

    const token = await calendarAuth.getAccessToken();
    await deleteTodoTask(token, listId, taskId);

    toolLogger?.info(`unified-inbox: tool deleted task ${taskId}`);
    return json({ success: true, taskId });
  },
};
