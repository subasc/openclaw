import { Type } from "@sinclair/typebox";
import fs from "node:fs/promises";
import path from "node:path";
import type { AgentBinding } from "../../config/types.js";
import { applyAgentBindings } from "../../commands/agents.bindings.js";
import {
  applyAgentConfig,
  buildAgentSummaries,
  findAgentEntryIndex,
  listAgentEntries,
  pruneAgentConfig,
} from "../../commands/agents.config.js";
import { loadConfig, writeConfigFile } from "../../config/config.js";
import { DEFAULT_AGENT_ID, normalizeAgentId } from "../../routing/session-key.js";
import { resolveAgentWorkspaceDir } from "../agent-scope.js";
import { stringEnum } from "../schema/typebox.js";
import { ensureAgentWorkspace } from "../workspace.js";
import { type AnyAgentTool, jsonResult, readStringParam } from "./common.js";

const AGENTS_MANAGE_ACTIONS = ["list", "create", "update", "delete", "exists"] as const;

const AgentsManageToolSchema = Type.Object({
  action: stringEnum(AGENTS_MANAGE_ACTIONS),
  agentId: Type.Optional(Type.String()),
  name: Type.Optional(Type.String()),
  model: Type.Optional(Type.String()),
  workspace: Type.Optional(Type.String()),
  identity: Type.Optional(Type.String()),
  identityName: Type.Optional(Type.String()),
  identityEmoji: Type.Optional(Type.String()),
  bindChannel: Type.Optional(Type.String()),
  bindPeerKind: Type.Optional(Type.String()),
  bindPeerId: Type.Optional(Type.String()),
  deleteFiles: Type.Optional(Type.Boolean()),
});

type AgentsManageToolOptions = {
  agentSessionKey?: string;
};

export function createAgentsManageTool(_opts?: AgentsManageToolOptions): AnyAgentTool {
  return {
    label: "Agents Manage",
    name: "agents_manage",
    description: `Create, update, delete, and inspect agents at runtime.

ACTIONS:
- list: List all configured agents with id, name, model, workspace, bindings
- exists: Check if an agent exists (requires agentId)
- create: Create a new agent (requires agentId; optional: name, model, workspace, identity, identityName, identityEmoji, bindChannel/bindPeerKind/bindPeerId)
- update: Update an existing agent's name, model, or workspace (requires agentId)
- delete: Remove an agent from config and optionally delete workspace files (requires agentId; deleteFiles defaults to true)

CONSTRAINTS:
- Cannot create or delete the "main" agent
- Agent IDs are normalized (lowercase, alphanumeric + hyphens)
- Duplicate agent IDs are rejected on create`,
    parameters: AgentsManageToolSchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      switch (action) {
        case "list": {
          const cfg = loadConfig();
          const summaries = buildAgentSummaries(cfg);
          return jsonResult({
            status: "ok",
            action: "list",
            agents: summaries.map((s) => ({
              id: s.id,
              name: s.name,
              identityName: s.identityName,
              identityEmoji: s.identityEmoji,
              workspace: s.workspace,
              model: s.model,
              bindings: s.bindings,
              isDefault: s.isDefault,
            })),
          });
        }

        case "exists": {
          const agentId = readStringParam(params, "agentId", { required: true });
          const id = normalizeAgentId(agentId);
          const cfg = loadConfig();
          const list = listAgentEntries(cfg);
          const index = findAgentEntryIndex(list, id);
          return jsonResult({
            status: "ok",
            action: "exists",
            agentId: id,
            exists: index >= 0,
          });
        }

        case "create": {
          const agentId = readStringParam(params, "agentId", { required: true });
          const id = normalizeAgentId(agentId);

          if (id === DEFAULT_AGENT_ID) {
            throw new Error("Cannot create the main agent.");
          }

          let cfg = loadConfig();
          const existingList = listAgentEntries(cfg);
          if (findAgentEntryIndex(existingList, id) >= 0) {
            throw new Error(`Agent "${id}" already exists.`);
          }

          const name = readStringParam(params, "name");
          const model = readStringParam(params, "model");
          const workspaceParam = readStringParam(params, "workspace");
          const identityContent = readStringParam(params, "identity");
          const identityName = readStringParam(params, "identityName");
          const identityEmoji = readStringParam(params, "identityEmoji");

          // Apply agent config
          cfg = applyAgentConfig(cfg, {
            agentId: id,
            name,
            model,
            workspace: workspaceParam,
          });

          // Ensure workspace exists
          const workspaceDir = workspaceParam ?? resolveAgentWorkspaceDir(cfg, id);
          await ensureAgentWorkspace({ dir: workspaceDir });

          // Write SOUL.md if identity content provided
          if (identityContent) {
            const soulPath = path.join(workspaceDir, "SOUL.md");
            await fs.writeFile(soulPath, identityContent, "utf-8");
          }

          // Write IDENTITY.md if name or emoji provided
          if (identityName || identityEmoji) {
            const identityPath = path.join(workspaceDir, "IDENTITY.md");
            const lines: string[] = [];
            if (identityName) {
              lines.push(`# ${identityName}`);
            }
            if (identityEmoji) {
              lines.push(`Emoji: :${identityEmoji}:`);
            }
            await fs.writeFile(identityPath, lines.join("\n\n") + "\n", "utf-8");
          }

          // Apply channel binding if provided
          const bindChannel = readStringParam(params, "bindChannel");
          const bindPeerKind = readStringParam(params, "bindPeerKind");
          const bindPeerId = readStringParam(params, "bindPeerId");
          if (bindChannel) {
            const match: AgentBinding["match"] = { channel: bindChannel };
            if (bindPeerKind && bindPeerId) {
              match.peer = { kind: bindPeerKind as "group" | "direct", id: bindPeerId };
            }
            const bindingResult = applyAgentBindings(cfg, [{ agentId: id, match }]);
            cfg = bindingResult.config;
          }

          await writeConfigFile(cfg);

          return jsonResult({
            status: "ok",
            action: "create",
            agentId: id,
            name: name ?? id,
            workspace: workspaceDir,
            text: `Agent "${name ?? id}" created successfully.`,
          });
        }

        case "update": {
          const agentId = readStringParam(params, "agentId", { required: true });
          const id = normalizeAgentId(agentId);

          let cfg = loadConfig();
          const existingList = listAgentEntries(cfg);
          if (findAgentEntryIndex(existingList, id) < 0) {
            throw new Error(`Agent "${id}" not found.`);
          }

          const name = readStringParam(params, "name");
          const model = readStringParam(params, "model");
          const workspaceParam = readStringParam(params, "workspace");

          if (!name && !model && !workspaceParam) {
            throw new Error("At least one of name, model, or workspace must be provided.");
          }

          cfg = applyAgentConfig(cfg, {
            agentId: id,
            name,
            model,
            workspace: workspaceParam,
          });

          await writeConfigFile(cfg);

          return jsonResult({
            status: "ok",
            action: "update",
            agentId: id,
            text: `Agent "${id}" updated successfully.`,
          });
        }

        case "delete": {
          const agentId = readStringParam(params, "agentId", { required: true });
          const id = normalizeAgentId(agentId);

          if (id === DEFAULT_AGENT_ID) {
            throw new Error("Cannot delete the main agent.");
          }

          const cfg = loadConfig();
          const existingList = listAgentEntries(cfg);
          if (findAgentEntryIndex(existingList, id) < 0) {
            throw new Error(`Agent "${id}" not found.`);
          }

          const deleteFiles = params.deleteFiles !== false;
          const workspaceDir = resolveAgentWorkspaceDir(cfg, id);

          const pruned = pruneAgentConfig(cfg, id);
          await writeConfigFile(pruned.config);

          if (deleteFiles) {
            try {
              await fs.rm(workspaceDir, { recursive: true, force: true });
            } catch {
              // Workspace deletion is best-effort
            }
          }

          return jsonResult({
            status: "ok",
            action: "delete",
            agentId: id,
            removedBindings: pruned.removedBindings,
            filesDeleted: deleteFiles,
            text: `Agent "${id}" deleted successfully.`,
          });
        }

        default:
          throw new Error(`Unknown action: ${action}`);
      }
    },
  };
}
