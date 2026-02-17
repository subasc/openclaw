// ============================================================================
// Agent tools: email & Teams actions the AI can invoke from Telegram
// ============================================================================

import { Type } from "@sinclair/typebox";
import type { IMsAuthProvider } from "./types.js";
import { sendMailViaOutlookRest, replyToEmailViaOutlookRest, sendChatMessage, listChats } from "./ms-graph-client.js";

type Logger = { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void };

// Shared auth providers — set by service.ts after startup
let mailAuth: IMsAuthProvider | null = null;
let chatAuth: IMsAuthProvider | null = null;
let toolLogger: Logger | null = null;

export function setToolAuthProviders(deps: {
  mailAuth: IMsAuthProvider;
  chatAuth: IMsAuthProvider;
  log: Logger;
}): void {
  mailAuth = deps.mailAuth;
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
    if (!mailAuth) throw new Error("Email auth not available — unified-inbox service may not be running");
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
    messageId: Type.String({ description: "The Microsoft Graph message ID of the email to reply to" }),
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
            preview: (c.lastMessagePreview.body?.content ?? "").replace(/<[^>]+>/g, "").slice(0, 120),
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
