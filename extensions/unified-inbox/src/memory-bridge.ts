/**
 * Memory Bridge — persists Unified Inbox notifications to LanceDB for long-term recall.
 *
 * Connects to the SAME LanceDB database and table used by memory-lancedb so the
 * agent's `memory_recall` tool and `before_agent_start` auto-recall can surface
 * old notifications that have already fallen out of the 24-hour ring buffer.
 *
 * Design:
 * - Lazy init: DB + OpenAI client created on first write
 * - Best-effort: all errors caught silently, never blocks notification delivery
 * - 95% cosine-similarity dedup (matches memory-lancedb threshold)
 * - Source-based importance scoring
 */

import { createHash, randomUUID } from "node:crypto";
import { homedir } from "node:os";
import { join } from "node:path";

// ---------------------------------------------------------------------------
// Constants (must match memory-lancedb exactly)
// ---------------------------------------------------------------------------

const TABLE_NAME = "memories";
const VECTOR_DIM = 1536; // text-embedding-3-small
const EMBEDDING_MODEL = "text-embedding-3-small";
const DB_PATH_SEGMENTS = [".openclaw", "memory", "lancedb"];
const DEDUP_MIN_SCORE = 0.95;

const IMPORTANCE_BY_SOURCE: Record<string, number> = {
  email: 0.6,
  calendar: 0.4,
  teams: 0.5,
  whatsapp: 0.4,
};

const CATEGORY_BY_SOURCE: Record<string, string> = {
  email: "entity",
  calendar: "fact",
  teams: "entity",
  whatsapp: "entity",
};

// ---------------------------------------------------------------------------
// In-memory dedup cache — catches rapid-fire duplicates before the vector
// index has time to reflect a just-added row (write-read consistency gap).
// ---------------------------------------------------------------------------

const RECENT_HASHES = new Set<string>();
const RECENT_HASHES_MAX = 200;

function textHash(text: string): string {
  return createHash("sha256").update(text).digest("hex");
}

function addToRecentHashes(hash: string): void {
  RECENT_HASHES.add(hash);
  if (RECENT_HASHES.size > RECENT_HASHES_MAX) {
    // Evict oldest (first inserted)
    const first = RECENT_HASHES.values().next().value;
    if (first) RECENT_HASHES.delete(first);
  }
}

// ---------------------------------------------------------------------------
// Lazy state
// ---------------------------------------------------------------------------

let lanceTable: import("@lancedb/lancedb").Table | null = null;
let openaiClient: import("openai").default | null = null;
let initPromise: Promise<boolean> | null = null;

function resolveDbPath(): string {
  return join(homedir(), ...DB_PATH_SEGMENTS);
}

function resolveApiKey(): string {
  const key = process.env.OPENAI_API_KEY;
  if (!key) throw new Error("memory-bridge: OPENAI_API_KEY not set");
  return key;
}

async function ensureInitialized(): Promise<boolean> {
  if (lanceTable && openaiClient) return true;
  if (initPromise) return initPromise;

  initPromise = doInit();
  return initPromise;
}

async function doInit(): Promise<boolean> {
  const lancedb = await import("@lancedb/lancedb");
  const OpenAI = (await import("openai")).default;

  const db = await lancedb.connect(resolveDbPath());
  const tables = await db.tableNames();

  if (tables.includes(TABLE_NAME)) {
    lanceTable = await db.openTable(TABLE_NAME);
  } else {
    // Create the table with the same schema as memory-lancedb
    lanceTable = await db.createTable(TABLE_NAME, [
      {
        id: "__schema__",
        text: "",
        vector: Array.from({ length: VECTOR_DIM }).fill(0),
        importance: 0,
        category: "other",
        createdAt: 0,
      },
    ]);
    await lanceTable.delete('id = "__schema__"');
  }

  openaiClient = new OpenAI({ apiKey: resolveApiKey() });
  return true;
}

async function embed(text: string): Promise<number[]> {
  const resp = await openaiClient!.embeddings.create({
    model: EMBEDDING_MODEL,
    input: text,
  });
  return resp.data[0].embedding;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Persist a notification summary to long-term vector memory.
 * Fire-and-forget — caller should `.catch(() => {})`.
 */
export async function persistToLongTermMemory(text: string, source: string): Promise<void> {
  try {
    await ensureInitialized();
  } catch {
    return; // LanceDB or OpenAI unavailable — silently skip
  }

  try {
    // Fast exact-text dedup (catches rapid-fire identical notifications)
    const hash = textHash(text);
    if (RECENT_HASHES.has(hash)) return;

    const vector = await embed(text);

    // Semantic dedup: skip if a near-identical memory already exists in LanceDB
    const existing = await lanceTable!.vectorSearch(vector).limit(1).toArray();
    if (existing.length > 0) {
      const distance = existing[0]._distance ?? 0;
      const score = 1 / (1 + distance);
      if (score >= DEDUP_MIN_SCORE) return;
    }

    await lanceTable!.add([
      {
        id: randomUUID(),
        text,
        vector,
        importance: IMPORTANCE_BY_SOURCE[source] ?? 0.5,
        category: CATEGORY_BY_SOURCE[source] ?? "other",
        createdAt: Date.now(),
      },
    ]);

    // Track hash only after successful write
    addToRecentHashes(hash);
  } catch {
    // Best-effort — never block notification delivery
  }
}
