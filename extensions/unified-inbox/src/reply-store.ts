// ============================================================================
// Persistent reply mapping store (telegramMsgId -> source context)
// JSON file with TTL-based cleanup
// ============================================================================

import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import type { ReplyMapping, ReplySourceContext, ReplySource } from "./types.js";

export type ReplyStoreOptions = {
  storeFile: string;
  maxEntries: number;
  ttlMs: number;
};

export class ReplyStore {
  private mappings: Map<number, ReplyMapping> = new Map();
  private filePath: string;
  private maxEntries: number;
  private ttlMs: number;
  private dirty = false;
  private flushTimer: ReturnType<typeof setInterval> | null = null;

  constructor(opts: ReplyStoreOptions) {
    this.filePath = resolveHomePath(opts.storeFile);
    this.maxEntries = opts.maxEntries;
    this.ttlMs = opts.ttlMs;
  }

  /** Load existing mappings from disk */
  async load(): Promise<void> {
    try {
      const raw = await fs.readFile(this.filePath, "utf-8");
      const entries = JSON.parse(raw) as ReplyMapping[];
      const now = Date.now();
      for (const entry of entries) {
        // Skip expired entries
        if (now - entry.createdAt < this.ttlMs) {
          this.mappings.set(entry.telegramMessageId, entry);
        }
      }
    } catch {
      // File doesn't exist or is invalid — start fresh
    }
  }

  /** Start periodic flush to disk (every 30 seconds if dirty) */
  startAutoFlush(): void {
    this.flushTimer = setInterval(async () => {
      if (this.dirty) {
        await this.flush();
      }
    }, 30_000);
  }

  /** Stop auto-flush */
  stopAutoFlush(): void {
    if (this.flushTimer) {
      clearInterval(this.flushTimer);
      this.flushTimer = null;
    }
  }

  /** Store a reply mapping */
  set(
    telegramMessageId: number,
    source: ReplySource,
    sourceContext: ReplySourceContext,
  ): void {
    this.mappings.set(telegramMessageId, {
      telegramMessageId,
      source,
      sourceContext,
      createdAt: Date.now(),
    });
    this.dirty = true;
    this.prune();
  }

  /** Look up a reply mapping by Telegram message ID */
  get(telegramMessageId: number): ReplyMapping | undefined {
    const entry = this.mappings.get(telegramMessageId);
    if (!entry) return undefined;
    // Check TTL
    if (Date.now() - entry.createdAt >= this.ttlMs) {
      this.mappings.delete(telegramMessageId);
      this.dirty = true;
      return undefined;
    }
    return entry;
  }

  /** Get the count of stored mappings */
  get size(): number {
    return this.mappings.size;
  }

  /** Flush to disk */
  async flush(): Promise<void> {
    const entries = Array.from(this.mappings.values());
    const dir = path.dirname(this.filePath);
    await fs.mkdir(dir, { recursive: true });
    await fs.writeFile(this.filePath, JSON.stringify(entries, null, 2));
    this.dirty = false;
  }

  /** Remove expired and excess entries */
  private prune(): void {
    const now = Date.now();

    // Remove expired
    for (const [id, entry] of this.mappings) {
      if (now - entry.createdAt >= this.ttlMs) {
        this.mappings.delete(id);
      }
    }

    // Remove oldest if over max
    if (this.mappings.size > this.maxEntries) {
      const sorted = Array.from(this.mappings.entries()).sort(
        (a, b) => a[1].createdAt - b[1].createdAt,
      );
      const toRemove = sorted.slice(0, sorted.length - this.maxEntries);
      for (const [id] of toRemove) {
        this.mappings.delete(id);
      }
    }
  }
}

function resolveHomePath(p: string): string {
  if (p.startsWith("~/")) {
    return path.join(os.homedir(), p.slice(2));
  }
  return p;
}
