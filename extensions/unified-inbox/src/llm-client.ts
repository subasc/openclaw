// ============================================================================
// Lightweight OpenAI API client for email summarization and draft generation
// Uses gpt-4o-mini for fast, cheap processing (~$0.001/email)
// ============================================================================

const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
const MODEL = "gpt-4o-mini";

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

function getApiKey(): string | undefined {
  return process.env.OPENAI_API_KEY;
}

async function chatCompletion(
  systemPrompt: string,
  userContent: string,
  log?: Logger,
): Promise<string | undefined> {
  const apiKey = getApiKey();
  if (!apiKey) {
    log?.warn("unified-inbox: OPENAI_API_KEY not set, skipping LLM call");
    return undefined;
  }

  try {
    const res = await fetch(OPENAI_API_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: MODEL,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        max_tokens: 500,
        temperature: 0.3,
      }),
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "");
      log?.error(`unified-inbox: OpenAI API ${res.status}: ${text.slice(0, 200)}`);
      return undefined;
    }

    const data = (await res.json()) as {
      choices?: Array<{ message?: { content?: string } }>;
    };

    return data.choices?.[0]?.message?.content?.trim();
  } catch (err) {
    log?.error(
      `unified-inbox: OpenAI API error: ${err instanceof Error ? err.message : String(err)}`,
    );
    return undefined;
  }
}

/**
 * Summarize an email for a mobile notification.
 * Returns plain text summary or undefined on failure (caller should fall back to bodyPreview).
 */
export async function summarizeEmail(
  params: { from: string; subject: string; body: string },
  log?: Logger,
): Promise<string | undefined> {
  const systemPrompt =
    "Summarize this email concisely. Start with a one-line summary, then list key points as bullet points using the bullet character. Keep it brief — this is for a mobile notification. Output plain text only, no markdown formatting.";

  const userContent = `From: ${params.from}\nSubject: ${params.subject}\n\n${params.body}`;

  return chatCompletion(systemPrompt, userContent, log);
}

/**
 * Draft a professional email reply from user's shorthand notes.
 * Returns the email body text or undefined on failure.
 */
export async function draftReply(
  params: {
    fromName: string;
    fromAddress: string;
    subject: string;
    bodyPreview: string;
    mode: "reply" | "reply-all";
    notes: string;
  },
  log?: Logger,
): Promise<string | undefined> {
  const modeLabel = params.mode === "reply-all" ? "reply-all" : "reply";

  const systemPrompt = [
    `Draft a professional ${modeLabel} to this email.`,
    "Output ONLY the email body text. Professional tone, proper grammar, no typos.",
    "Do not include subject line, greeting headers like 'Subject:', or email metadata.",
    "Just write the reply body that will be sent.",
  ].join("\n");

  const userContent = [
    `Original from: ${params.fromName} <${params.fromAddress}>`,
    `Subject: ${params.subject}`,
    `Preview: ${params.bodyPreview}`,
    `Mode: ${modeLabel}`,
    `User's notes/bullet points: ${params.notes}`,
  ].join("\n");

  return chatCompletion(systemPrompt, userContent, log);
}
