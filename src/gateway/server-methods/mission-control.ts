import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { GatewayRequestHandlers } from "./types.js";
import { resolveDefaultAgentId } from "../../agents/agent-scope.js";
import { loadConfig } from "../../config/config.js";
import { normalizeAgentId } from "../../routing/session-key.js";
import { ErrorCodes, errorShape } from "../protocol/index.js";
import { formatForLog } from "../ws-log.js";

function getCpuUsagePercent(): number {
  const cpus = os.cpus();
  if (cpus.length === 0) {
    return 0;
  }
  let totalIdle = 0;
  let totalTick = 0;
  for (const cpu of cpus) {
    const { user, nice, sys, idle, irq } = cpu.times;
    totalTick += user + nice + sys + idle + irq;
    totalIdle += idle;
  }
  return Math.round(((totalTick - totalIdle) / totalTick) * 100);
}

function getLanceDbSizeBytes(): {
  available: boolean;
  dbPath: string | null;
  dbSizeBytes: number | null;
} {
  const home = os.homedir();
  const dbPath = path.join(home, ".openclaw", "memory", "lancedb");
  try {
    const stat = fs.statSync(dbPath);
    if (stat.isDirectory()) {
      let totalSize = 0;
      const entries = fs.readdirSync(dbPath, { withFileTypes: true, recursive: true });
      for (const entry of entries) {
        if (entry.isFile()) {
          try {
            const parentDir = entry.parentPath ?? entry.path ?? dbPath;
            const filePath = path.join(parentDir, entry.name);
            const fileStat = fs.statSync(filePath);
            totalSize += fileStat.size;
          } catch {
            // skip unreadable files
          }
        }
      }
      return { available: true, dbPath, dbSizeBytes: totalSize };
    }
    return { available: true, dbPath, dbSizeBytes: stat.size };
  } catch {
    return { available: false, dbPath, dbSizeBytes: null };
  }
}

type AgentModelInfo = {
  agentId: string;
  name?: string;
  isDefault: boolean;
  model: string | null;
  modelProvider?: string;
  fallbacks?: string[];
};

function resolveAgentModelInfo(): AgentModelInfo[] {
  try {
    const cfg = loadConfig();
    const defaultId = normalizeAgentId(resolveDefaultAgentId(cfg));
    const agents: AgentModelInfo[] = [];

    // Global default model
    const globalModel = cfg.agents?.defaults?.model;
    const globalModelPrimary =
      typeof globalModel === "string"
        ? globalModel
        : typeof globalModel === "object" && globalModel !== null
          ? ((globalModel as { primary?: string }).primary ?? null)
          : null;
    const globalFallbacks =
      typeof globalModel === "object" && globalModel !== null
        ? ((globalModel as { fallbacks?: string[] }).fallbacks ?? [])
        : [];

    const agentList = cfg.agents?.list ?? [];
    const seenIds = new Set<string>();

    for (const entry of agentList) {
      if (!entry?.id) {
        continue;
      }
      const id = normalizeAgentId(entry.id);
      seenIds.add(id);

      const agentModel = entry.model;
      let primary: string | null = null;
      let fallbacks: string[] = [];

      if (typeof agentModel === "string") {
        primary = agentModel;
      } else if (typeof agentModel === "object" && agentModel !== null) {
        primary = (agentModel as { primary?: string }).primary ?? null;
        fallbacks = (agentModel as { fallbacks?: string[] }).fallbacks ?? [];
      }

      agents.push({
        agentId: id,
        name: entry.name ?? entry.identity?.name ?? undefined,
        isDefault: id === defaultId,
        model: primary || globalModelPrimary,
        fallbacks:
          fallbacks.length > 0
            ? fallbacks
            : globalFallbacks.length > 0
              ? globalFallbacks
              : undefined,
      });
    }

    // Ensure default agent is included even if not in list
    if (!seenIds.has(defaultId)) {
      agents.unshift({
        agentId: defaultId,
        isDefault: true,
        model: globalModelPrimary,
        fallbacks: globalFallbacks.length > 0 ? globalFallbacks : undefined,
      });
    }

    return agents;
  } catch {
    return [];
  }
}

export const missionControlHandlers: GatewayRequestHandlers = {
  "mission-control.snapshot": async ({ respond, context }) => {
    try {
      const { getHealthCache, refreshHealthSnapshot } = context;

      // Gather health data (use cache if available, refresh in background)
      let health = getHealthCache();
      if (!health) {
        try {
          health = await refreshHealthSnapshot({ probe: false });
        } catch {
          health = null;
        }
      }

      // System metrics
      const cpuUsagePercent = getCpuUsagePercent();
      const cpuCores = os.cpus().length;
      const memTotalBytes = os.totalmem();
      const memFreeBytes = os.freemem();
      const memUsedPercent = Math.round(((memTotalBytes - memFreeBytes) / memTotalBytes) * 100);
      const processMemMB = Math.round((process.memoryUsage().rss / 1024 / 1024) * 10) / 10;
      const uptimeMs = health?.ts ? Date.now() - health.ts + (health.durationMs ?? 0) : null;

      // LanceDB / LTM stats
      const memory = getLanceDbSizeBytes();

      // Per-agent model configuration
      const agentModels = resolveAgentModelInfo();

      const snapshot = {
        ts: Date.now(),
        system: {
          cpuUsagePercent,
          cpuCores,
          memTotalBytes,
          memFreeBytes,
          memUsedPercent,
          processMemMB,
          uptimeMs,
          platform: process.platform,
          nodeVersion: process.version,
        },
        health: health ?? null,
        memory,
        agentModels,
      };

      respond(true, snapshot, undefined);
    } catch (err) {
      respond(false, undefined, errorShape(ErrorCodes.UNAVAILABLE, formatForLog(err)));
    }
  },
};
