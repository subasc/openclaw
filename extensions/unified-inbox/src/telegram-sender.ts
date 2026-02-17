// ============================================================================
// Direct Telegram Bot API client (decoupled from Telegram extension)
// ============================================================================

const TELEGRAM_API = "https://api.telegram.org";

export type TelegramSendResult = {
  ok: boolean;
  messageId?: number;
  error?: string;
};

/**
 * Send a text message via the Telegram Bot API.
 * Returns the message_id of the sent message (needed for reply tracking).
 */
export async function sendTelegramMessage(params: {
  botToken: string;
  chatId: string;
  text: string;
  parseMode?: "MarkdownV2" | "HTML";
  replyToMessageId?: number;
  replyMarkup?: { inline_keyboard: Array<Array<{ text: string; callback_data: string }>> };
}): Promise<TelegramSendResult> {
  const { botToken, chatId, text, parseMode, replyToMessageId, replyMarkup } = params;

  const body: Record<string, unknown> = {
    chat_id: chatId,
    text,
  };

  if (parseMode) body.parse_mode = parseMode;
  if (replyToMessageId) {
    body.reply_parameters = { message_id: replyToMessageId };
  }
  if (replyMarkup) body.reply_markup = replyMarkup;

  try {
    const res = await fetch(`${TELEGRAM_API}/bot${botToken}/sendMessage`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    const data = (await res.json()) as {
      ok: boolean;
      result?: { message_id: number };
      description?: string;
    };

    if (data.ok && data.result) {
      return { ok: true, messageId: data.result.message_id };
    }

    // If MarkdownV2 parse fails, retry as plain text
    if (!data.ok && parseMode) {
      return sendTelegramMessage({ ...params, parseMode: undefined });
    }

    return { ok: false, error: data.description ?? "Unknown Telegram error" };
  } catch (err) {
    return {
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

/**
 * Edit an existing Telegram message text (and optionally update inline keyboard).
 */
export async function editTelegramMessage(params: {
  botToken: string;
  chatId: string;
  messageId: number;
  text: string;
  parseMode?: "MarkdownV2" | "HTML";
  replyMarkup?: { inline_keyboard: Array<Array<{ text: string; callback_data: string }>> };
}): Promise<TelegramSendResult> {
  const { botToken, chatId, messageId, text, parseMode, replyMarkup } = params;

  const body: Record<string, unknown> = {
    chat_id: chatId,
    message_id: messageId,
    text,
  };

  if (parseMode) body.parse_mode = parseMode;
  if (replyMarkup) body.reply_markup = replyMarkup;

  try {
    const res = await fetch(`${TELEGRAM_API}/bot${botToken}/editMessageText`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    const data = (await res.json()) as {
      ok: boolean;
      result?: { message_id: number };
      description?: string;
    };

    if (data.ok && data.result) {
      return { ok: true, messageId: data.result.message_id };
    }

    // If MarkdownV2 parse fails, retry as plain text
    if (!data.ok && parseMode) {
      return editTelegramMessage({ ...params, parseMode: undefined });
    }

    return { ok: false, error: data.description ?? "Unknown Telegram error" };
  } catch (err) {
    return {
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}
