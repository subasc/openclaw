// ============================================================================
// Sub-Bot Registry — wraps monitors as named, manageable Sub Bots
// ============================================================================

import type { MonitorStatus } from "./types.js";

export type SubBotType = "email" | "calendar" | "teams-chat" | "whatsapp";

export interface SubBot {
  readonly id: string;
  readonly name: string;
  readonly description: string;
  readonly type: SubBotType;
  readonly status: MonitorStatus;
  start(): Promise<void>;
  stop(): void;
  restart(): Promise<void>;
}

type Logger = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
};

// ── Registry ────────────────────────────────────────────────────────────────

export class SubBotRegistry {
  private bots = new Map<string, SubBot>();

  constructor(private log: Logger) {}

  register(bot: SubBot): void {
    this.bots.set(bot.id, bot);
    this.log.info(`sub-bot-registry: registered "${bot.name}" [${bot.id}]`);
  }

  get(id: string): SubBot | undefined {
    return this.bots.get(id);
  }

  getAll(): SubBot[] {
    return [...this.bots.values()];
  }

  async startBot(id: string): Promise<string> {
    const bot = this.bots.get(id);
    if (!bot) return `Unknown bot: ${id}. Available: ${this.availableIds()}`;
    if (bot.status.running) return `${bot.name} is already running.`;
    try {
      await bot.start();
      return `${bot.name} started.`;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.log.error(`sub-bot-registry: failed to start ${id}: ${msg}`);
      return `Failed to start ${bot.name}: ${msg}`;
    }
  }

  stopBot(id: string): string {
    const bot = this.bots.get(id);
    if (!bot) return `Unknown bot: ${id}. Available: ${this.availableIds()}`;
    if (!bot.status.running) return `${bot.name} is already stopped.`;
    try {
      bot.stop();
      return `${bot.name} stopped.`;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.log.error(`sub-bot-registry: failed to stop ${id}: ${msg}`);
      return `Failed to stop ${bot.name}: ${msg}`;
    }
  }

  async restartBot(id: string): Promise<string> {
    const bot = this.bots.get(id);
    if (!bot) return `Unknown bot: ${id}. Available: ${this.availableIds()}`;
    try {
      await bot.restart();
      return `${bot.name} restarted.`;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.log.error(`sub-bot-registry: failed to restart ${id}: ${msg}`);
      return `Failed to restart ${bot.name}: ${msg}`;
    }
  }

  formatStatusSummary(): string {
    if (this.bots.size === 0) return "No sub-bots registered.";

    const lines: string[] = [];
    for (const bot of this.bots.values()) {
      const state = bot.status.paused
        ? "PAUSED"
        : bot.status.running
          ? "running"
          : "stopped";
      const lastPoll = bot.status.lastPollAt
        ? new Date(bot.status.lastPollAt).toLocaleTimeString()
        : "n/a";
      const error = bot.status.lastError
        ? `\n  Last error: ${bot.status.lastError.slice(0, 80)}`
        : "";
      lines.push(`${bot.name} [${bot.id}]: ${state} (last poll: ${lastPoll})${error}`);
    }
    return lines.join("\n");
  }

  private availableIds(): string {
    return [...this.bots.keys()].join(", ");
  }
}

// ── Factories ───────────────────────────────────────────────────────────────

/** Wraps a polling monitor (EmailMonitor, CalendarMonitor, TeamsChatMonitor) as a SubBot. */
export function createMonitorSubBot(opts: {
  id: string;
  name: string;
  description: string;
  type: SubBotType;
  monitor: { readonly status: MonitorStatus; start(): Promise<void>; stop(): void };
}): SubBot {
  const { id, name, description, type, monitor } = opts;
  return {
    id,
    name,
    description,
    type,
    get status() {
      return monitor.status;
    },
    start: () => monitor.start(),
    stop: () => monitor.stop(),
    async restart() {
      monitor.stop();
      await monitor.start();
    },
  };
}

/** Wraps the WhatsApp bridge as a SubBot with synthetic MonitorStatus. */
export function createWhatsAppSubBot(opts: {
  id?: string;
  name?: string;
  description?: string;
}): SubBot {
  const id = opts.id ?? "whatsapp";
  const name = opts.name ?? "WhatsApp Bridge";
  const description = opts.description ?? "Forwards WhatsApp messages to Telegram";

  const status: MonitorStatus = {
    running: true,
    lastPollAt: null,
    lastErrorAt: null,
    lastError: null,
    consecutiveFailures: 0,
    paused: false,
  };

  return {
    id,
    name,
    description,
    type: "whatsapp",
    status,
    async start() {
      status.running = true;
      status.paused = false;
    },
    stop() {
      status.running = false;
    },
    async restart() {
      status.running = true;
      status.paused = false;
    },
  };
}
