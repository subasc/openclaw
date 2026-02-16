// ============================================================================
// Configuration schema and resolution for the Unified Inbox extension
// ============================================================================

/** Unified Inbox plugin configuration */
export type UnifiedInboxConfig = {
  enabled: boolean;
  authMode: "browser" | "device-code";
  browserProfile: string;
  browserCdpPort: number;
  telegramChatId: string;
  telegramBotToken: string;
  telegramAccountId: string;
  microsoft: {
    clientId: string;
    tenantId: string;
    tokenFile: string;
  };
  email: {
    enabled: boolean;
    pollIntervalMs: number;
    filterUnread: boolean;
    folders: string[];
    maxPerPoll: number;
  };
  calendar: {
    enabled: boolean;
    pollIntervalMs: number;
    lookAheadMinutes: number;
    reminderMinutes: number[];
  };
  teamsChat: {
    enabled: boolean;
    pollIntervalMs: number;
    excludeBotMessages: boolean;
  };
  whatsapp: {
    enabled: boolean;
    whatsappAccountId: string;
  };
  replyTracking: {
    storeFile: string;
    maxEntries: number;
    ttlMs: number;
  };
};

// Microsoft Office Desktop public client ID (no app registration needed)
const DEFAULT_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c";

const DEFAULTS: UnifiedInboxConfig = {
  enabled: false,
  authMode: "browser",
  browserProfile: "inbox",
  browserCdpPort: 18801,
  telegramChatId: "",
  telegramBotToken: "",
  telegramAccountId: "default",
  microsoft: {
    clientId: DEFAULT_CLIENT_ID,
    tenantId: "organizations",
    tokenFile: "~/.openclaw/unified-inbox-tokens.json",
  },
  email: {
    enabled: true,
    pollIntervalMs: 30_000,
    filterUnread: true,
    folders: ["Inbox"],
    maxPerPoll: 10,
  },
  calendar: {
    enabled: true,
    pollIntervalMs: 300_000,
    lookAheadMinutes: 60,
    reminderMinutes: [15, 5],
  },
  teamsChat: {
    enabled: true,
    pollIntervalMs: 15_000,
    excludeBotMessages: true,
  },
  whatsapp: {
    enabled: true,
    whatsappAccountId: "default",
  },
  replyTracking: {
    storeFile: "~/.openclaw/unified-inbox-replies.json",
    maxEntries: 10_000,
    ttlMs: 7 * 24 * 60 * 60 * 1000, // 7 days
  },
};

/** Resolve the full config with defaults applied */
export function resolveUnifiedInboxConfig(
  pluginConfig?: Record<string, unknown>,
): UnifiedInboxConfig {
  if (!pluginConfig) return DEFAULTS;

  const raw = pluginConfig as Partial<Record<string, unknown>>;

  return {
    enabled: Boolean(raw.enabled ?? DEFAULTS.enabled),
    authMode: (raw.authMode === "device-code" ? "device-code" : DEFAULTS.authMode),
    browserProfile: String(raw.browserProfile ?? DEFAULTS.browserProfile),
    browserCdpPort: Number(raw.browserCdpPort ?? DEFAULTS.browserCdpPort),
    telegramChatId: String(raw.telegramChatId ?? process.env.UNIFIED_INBOX_TELEGRAM_CHAT_ID ?? DEFAULTS.telegramChatId),
    telegramBotToken: String(raw.telegramBotToken ?? process.env.TELEGRAM_BOT_TOKEN ?? DEFAULTS.telegramBotToken),
    telegramAccountId: String(raw.telegramAccountId ?? DEFAULTS.telegramAccountId),
    microsoft: mergeSection(raw.microsoft, DEFAULTS.microsoft, {
      clientId: String,
      tenantId: String,
      tokenFile: String,
    }),
    email: mergeSection(raw.email, DEFAULTS.email, {
      enabled: Boolean,
      pollIntervalMs: Number,
      filterUnread: Boolean,
      folders: asStringArray,
      maxPerPoll: Number,
    }),
    calendar: mergeSection(raw.calendar, DEFAULTS.calendar, {
      enabled: Boolean,
      pollIntervalMs: Number,
      lookAheadMinutes: Number,
      reminderMinutes: asNumberArray,
    }),
    teamsChat: mergeSection(raw.teamsChat, DEFAULTS.teamsChat, {
      enabled: Boolean,
      pollIntervalMs: Number,
      excludeBotMessages: Boolean,
    }),
    whatsapp: mergeSection(raw.whatsapp, DEFAULTS.whatsapp, {
      enabled: Boolean,
      whatsappAccountId: String,
    }),
    replyTracking: mergeSection(raw.replyTracking, DEFAULTS.replyTracking, {
      storeFile: String,
      maxEntries: Number,
      ttlMs: Number,
    }),
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function mergeSection<T extends Record<string, unknown>>(
  raw: unknown,
  defaults: T,
  casters: Record<keyof T, (v: unknown) => unknown>,
): T {
  if (!raw || typeof raw !== "object") return defaults;
  const src = raw as Record<string, unknown>;
  const result = { ...defaults };
  for (const key of Object.keys(defaults)) {
    if (key in src && src[key] !== undefined) {
      (result as Record<string, unknown>)[key] = casters[key as keyof T](src[key]);
    }
  }
  return result;
}

function asStringArray(v: unknown): string[] {
  if (Array.isArray(v)) return v.map(String);
  return [];
}

function asNumberArray(v: unknown): number[] {
  if (Array.isArray(v)) return v.map(Number);
  return [];
}
