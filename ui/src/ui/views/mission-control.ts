import { html, nothing } from "lit";
import type { CostUsageSummary, MissionControlBridge, MissionControlSnapshot } from "../types.ts";

export type MissionControlProps = {
  loading: boolean;
  error: string | null;
  snapshot: MissionControlSnapshot | null;
  bridges: MissionControlBridge[] | null;
  costSummary: CostUsageSummary | null;
  lastRefresh: number | null;
  taskDraft: string;
  onRefresh: () => void;
  onTaskDraftChange: (value: string) => void;
  onCreateTask: () => void;
};

function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  if (bytes < 1024 * 1024 * 1024) {
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

function formatCost(cost: number): string {
  if (cost === 0) {
    return "$0.00";
  }
  if (cost < 0.01) {
    return `$${cost.toFixed(4)}`;
  }
  return `$${cost.toFixed(2)}`;
}

function timeAgo(ts: number | null): string {
  if (!ts) {
    return "never";
  }
  const seconds = Math.floor((Date.now() - ts) / 1000);
  if (seconds < 5) {
    return "just now";
  }
  if (seconds < 60) {
    return `${seconds}s ago`;
  }
  if (seconds < 3600) {
    return `${Math.floor(seconds / 60)}m ago`;
  }
  if (seconds < 86400) {
    return `${Math.floor(seconds / 3600)}h ago`;
  }
  return `${Math.floor(seconds / 86400)}d ago`;
}

function statusDotClass(running: boolean, hasError: boolean, paused: boolean): string {
  if (paused) {
    return "statusDot warn";
  }
  if (hasError) {
    return "statusDot danger";
  }
  if (running) {
    return "statusDot ok";
  }
  return "statusDot";
}

function bridgeStatusLabel(bridge: MissionControlBridge): string {
  if (bridge.status.paused) {
    return "paused";
  }
  if (!bridge.status.running) {
    return "stopped";
  }
  if (bridge.status.consecutiveFailures > 0) {
    return "degraded";
  }
  return "running";
}

function chipClass(running: boolean, hasError: boolean, paused: boolean): string {
  if (paused) {
    return "chip chip-warn";
  }
  if (hasError) {
    return "chip chip-danger";
  }
  if (running) {
    return "chip chip-ok";
  }
  return "chip";
}

export function renderMissionControl(props: MissionControlProps) {
  const snap = props.snapshot;
  const sys = snap?.system;
  const health = snap?.health as Record<string, unknown> | null;
  const mem = snap?.memory;
  const bridges = props.bridges;
  const agentModels = snap?.agentModels ?? [];
  const costSummary = props.costSummary;

  // Derive agent info from health snapshot
  const agents = health?.agents as
    | Array<{
        agentId: string;
        name?: string;
        isDefault?: boolean;
        heartbeat?: Record<string, unknown>;
      }>
    | undefined;
  const agentCount = agentModels.length || agents?.length || 0;

  // Bridge summary
  const bridgeCount = bridges?.length ?? 0;
  const bridgesOk =
    bridges?.filter((b) => b.status.running && b.status.consecutiveFailures === 0).length ?? 0;

  // Channel info from health snapshot
  const channelOrder = health?.channelOrder as string[] | undefined;
  const channelLabels = health?.channelLabels as Record<string, string> | undefined;
  const channels = health?.channels as Record<string, Record<string, unknown>> | undefined;

  // Cost summary
  const totalCost30d = costSummary?.totals?.totalCost ?? 0;
  const totalTokens30d = costSummary?.totals?.totalTokens ?? 0;

  return html`
    <section>
      <!-- Stat Row -->
      <div class="stat-grid">
        <div class="card" style="text-align: center;">
          <div class="muted" style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em;">Agents</div>
          <div style="font-size: 28px; font-weight: 700; margin: 4px 0;">${agentCount}</div>
          <div class="muted" style="font-size: 12px;">${agentModels.filter((a) => a.model).length} with model</div>
        </div>
        <div class="card" style="text-align: center;">
          <div class="muted" style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em;">Bridges</div>
          <div style="font-size: 28px; font-weight: 700; margin: 4px 0;">${bridgesOk}/${bridgeCount}</div>
          <div class="muted" style="font-size: 12px;">${bridgeCount - bridgesOk === 0 ? "all healthy" : `${bridgeCount - bridgesOk} issues`}</div>
        </div>
        <div class="card" style="text-align: center;">
          <div class="muted" style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em;">CPU</div>
          <div style="font-size: 28px; font-weight: 700; margin: 4px 0;">${sys?.cpuUsagePercent ?? "—"}%</div>
          <div class="muted" style="font-size: 12px;">${sys?.cpuCores ?? "—"} cores</div>
        </div>
        <div class="card" style="text-align: center;">
          <div class="muted" style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em;">Cost (30d)</div>
          <div style="font-size: 28px; font-weight: 700; margin: 4px 0;">${formatCost(totalCost30d)}</div>
          <div class="muted" style="font-size: 12px;">${totalTokens30d > 0 ? `${(totalTokens30d / 1000).toFixed(0)}K tokens` : "no usage"}</div>
        </div>
      </div>

      <!-- Agent Fleet + Bridge Status -->
      <div class="grid grid-cols-2" style="margin-top: 16px;">
        <div class="card">
          <div class="card-title">Agent Fleet</div>
          <div class="card-sub">Configured agents, their LLM models, and fallbacks.</div>
          <div class="stack" style="margin-top: 12px;">
            ${
              agentModels.length > 0
                ? agentModels.map(
                    (agent) => html`
                    <div style="padding: 8px 0; border-bottom: 1px solid var(--border);">
                      <div class="row" style="justify-content: space-between;">
                        <div style="display: flex; align-items: center; gap: 6px;">
                          <span style="font-weight: 600;">${agent.name || agent.agentId}</span>
                          ${
                            agent.isDefault
                              ? html`
                                  <span class="chip chip-ok" style="font-size: 10px">default</span>
                                `
                              : nothing
                          }
                        </div>
                      </div>
                      <div style="margin-top: 4px; display: flex; flex-wrap: wrap; gap: 4px; align-items: center;">
                        ${
                          agent.model
                            ? html`<span class="chip" style="font-size: 11px;">${agent.model}</span>`
                            : html`
                                <span class="muted" style="font-size: 11px">no model configured</span>
                              `
                        }
                        ${
                          agent.fallbacks && agent.fallbacks.length > 0
                            ? agent.fallbacks.map(
                                (fb) =>
                                  html`<span class="chip" style="font-size: 10px; opacity: 0.7;" title="fallback">${fb}</span>`,
                              )
                            : nothing
                        }
                      </div>
                    </div>
                  `,
                  )
                : agents && agents.length > 0
                  ? agents.map(
                      (agent) => html`
                      <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                        <div>
                          <span style="font-weight: 600;">${agent.name || agent.agentId}</span>
                          ${
                            agent.isDefault
                              ? html`
                                  <span class="chip chip-ok" style="margin-left: 6px; font-size: 10px">default</span>
                                `
                              : nothing
                          }
                        </div>
                        <span class="chip chip-ok" style="font-size: 11px;">configured</span>
                      </div>
                    `,
                    )
                  : html`
                      <div class="muted">No agents loaded yet.</div>
                    `
            }
          </div>
        </div>
        <div class="card">
          <div class="card-title">Bridge Status</div>
          <div class="card-sub">Unified Inbox monitors (email, calendar, Teams, WhatsApp).</div>
          <div class="stack" style="margin-top: 12px;">
            ${
              bridges && bridges.length > 0
                ? bridges.map(
                    (bridge) => html`
                    <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                      <div style="display: flex; align-items: center; gap: 8px;">
                        <span class="${statusDotClass(bridge.status.running, bridge.status.consecutiveFailures > 0, bridge.status.paused)}"></span>
                        <span style="font-weight: 500;">${bridge.name}</span>
                      </div>
                      <div style="display: flex; align-items: center; gap: 8px;">
                        <span class="${chipClass(bridge.status.running, bridge.status.consecutiveFailures > 0, bridge.status.paused)}" style="font-size: 11px;">
                          ${bridgeStatusLabel(bridge)}
                        </span>
                        <span class="muted" style="font-size: 11px;">${timeAgo(bridge.status.lastPollAt)}</span>
                      </div>
                    </div>
                  `,
                  )
                : html`
                    <div class="muted">No bridges configured or unified-inbox not loaded.</div>
                  `
            }
          </div>
        </div>
      </div>

      <!-- Channel Health + LTM / Vector Memory Stats -->
      <div class="grid grid-cols-2" style="margin-top: 16px;">
        <div class="card">
          <div class="card-title">Channel Health</div>
          <div class="card-sub">Connected messaging channels and their runtime status.</div>
          <div class="stack" style="margin-top: 12px;">
            ${
              channelOrder && channelOrder.length > 0
                ? channelOrder.map((channelId) => {
                    const label = channelLabels?.[channelId] ?? channelId;
                    const ch = channels?.[channelId];
                    const running = ch?.running === true;
                    const connected = ch?.connected === true;
                    const configured = ch?.configured === true;
                    const lastError = ch?.lastError as string | null | undefined;
                    const isOk = running || connected;
                    const hasError = Boolean(lastError);
                    return html`
                    <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                      <div style="display: flex; align-items: center; gap: 8px;">
                        <span class="${statusDotClass(isOk, hasError, false)}"></span>
                        <span style="font-weight: 500;">${label}</span>
                      </div>
                      <span class="muted" style="font-size: 11px;">
                        ${!configured ? "not configured" : isOk ? (connected ? "connected" : "running") : "stopped"}
                      </span>
                    </div>
                  `;
                  })
                : html`
                    <div class="muted">No channels available.</div>
                  `
            }
          </div>
        </div>
        <div class="card">
          <div class="card-title">Vector Memory (LTM)</div>
          <div class="card-sub">LanceDB long-term memory storage status.</div>
          <div class="stack" style="margin-top: 12px;">
            ${
              mem
                ? html`
                  <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                    <span class="muted">Status</span>
                    <span class="${mem.available ? "chip chip-ok" : "chip"}" style="font-size: 11px;">
                      ${mem.available ? "available" : "not found"}
                    </span>
                  </div>
                  <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                    <span class="muted">DB Size</span>
                    <span style="font-weight: 500;">${mem.dbSizeBytes != null ? formatBytes(mem.dbSizeBytes) : "—"}</span>
                  </div>
                  <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                    <span class="muted">Path</span>
                    <span class="mono" style="font-size: 11px; word-break: break-all;">${mem.dbPath ?? "—"}</span>
                  </div>
                  ${
                    sys
                      ? html`
                        <div class="row" style="justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--border);">
                          <span class="muted">Platform</span>
                          <span style="font-weight: 500;">${sys.platform}</span>
                        </div>
                        <div class="row" style="justify-content: space-between; padding: 6px 0;">
                          <span class="muted">Node.js</span>
                          <span class="mono" style="font-size: 11px;">${sys.nodeVersion}</span>
                        </div>
                      `
                      : nothing
                  }
                `
                : html`
                    <div class="muted">Loading memory stats...</div>
                  `
            }
          </div>
        </div>
      </div>

      <!-- LLM Cost Summary -->
      ${
        costSummary
          ? html`
            <div class="card" style="margin-top: 16px;">
              <div class="card-title">LLM Usage (Last 30 Days)</div>
              <div class="card-sub">Token usage and cost breakdown across all agents.</div>
              <div class="stat-grid" style="margin-top: 12px;">
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Total Cost</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${formatCost(costSummary.totals.totalCost)}</div>
                </div>
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Total Tokens</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${(costSummary.totals.totalTokens / 1000).toFixed(0)}K</div>
                </div>
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Input Cost</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${formatCost(costSummary.totals.inputCost ?? 0)}</div>
                </div>
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Output Cost</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${formatCost(costSummary.totals.outputCost ?? 0)}</div>
                </div>
              </div>
              ${
                costSummary.daily && costSummary.daily.length > 0
                  ? html`
                    <div style="margin-top: 12px; max-height: 180px; overflow-y: auto;">
                      <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                        <thead>
                          <tr style="border-bottom: 1px solid var(--border);">
                            <th style="text-align: left; padding: 4px 8px; font-weight: 600;">Date</th>
                            <th style="text-align: right; padding: 4px 8px; font-weight: 600;">Tokens</th>
                            <th style="text-align: right; padding: 4px 8px; font-weight: 600;">Cost</th>
                          </tr>
                        </thead>
                        <tbody>
                          ${costSummary.daily
                            .slice()
                            .toReversed()
                            .slice(0, 7)
                            .map(
                              (day) => html`
                                <tr style="border-bottom: 1px solid var(--border);">
                                  <td style="padding: 4px 8px;">${day.date}</td>
                                  <td style="text-align: right; padding: 4px 8px; font-family: var(--font-mono);">${(day.totalTokens / 1000).toFixed(0)}K</td>
                                  <td style="text-align: right; padding: 4px 8px; font-family: var(--font-mono);">${formatCost(day.totalCost)}</td>
                                </tr>
                              `,
                            )}
                        </tbody>
                      </table>
                    </div>
                  `
                  : nothing
              }
            </div>
          `
          : nothing
      }

      <!-- Quick Actions -->
      <div class="card" style="margin-top: 16px;">
        <div class="row" style="justify-content: space-between; align-items: center;">
          <div>
            <div class="card-title">Quick Actions</div>
            <div class="card-sub">
              ${props.lastRefresh ? html`Last refresh: ${timeAgo(props.lastRefresh)}` : "Not yet loaded."}
              ${props.error ? html` <span class="pill danger" style="margin-left: 8px;">${props.error}</span>` : nothing}
            </div>
          </div>
          <button class="btn" ?disabled=${props.loading} @click=${props.onRefresh}>
            ${props.loading ? "Refreshing..." : "Refresh All"}
          </button>
        </div>
        <div style="margin-top: 12px; display: flex; gap: 8px; align-items: center;">
          <input
            type="text"
            class="input"
            placeholder="Create a task..."
            .value=${props.taskDraft}
            @input=${(e: InputEvent) => props.onTaskDraftChange((e.target as HTMLInputElement).value)}
            @keydown=${(e: KeyboardEvent) => {
              if (e.key === "Enter" && props.taskDraft.trim()) {
                props.onCreateTask();
              }
            }}
            style="flex: 1;"
          />
          <button
            class="btn primary"
            ?disabled=${!props.taskDraft.trim()}
            @click=${props.onCreateTask}
          >
            + Task
          </button>
        </div>
      </div>

      <!-- System Info -->
      ${
        sys
          ? html`
            <div class="card" style="margin-top: 16px;">
              <div class="card-title">System Resources</div>
              <div class="card-sub">Host machine resource utilization.</div>
              <div class="stat-grid" style="margin-top: 12px;">
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">CPU Usage</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${sys.cpuUsagePercent}%</div>
                  <div class="muted" style="font-size: 11px;">${sys.cpuCores} cores</div>
                </div>
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Memory Used</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${formatBytes(sys.memTotalBytes - sys.memFreeBytes)}</div>
                  <div class="muted" style="font-size: 11px;">${sys.memUsedPercent}% of ${formatBytes(sys.memTotalBytes)}</div>
                </div>
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Process Memory</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${sys.processMemMB} MB</div>
                  <div class="muted" style="font-size: 11px;">gateway RSS</div>
                </div>
                <div style="text-align: center;">
                  <div class="muted" style="font-size: 11px; text-transform: uppercase;">Memory Free</div>
                  <div style="font-size: 22px; font-weight: 700; margin: 2px 0;">${formatBytes(sys.memFreeBytes)}</div>
                  <div class="muted" style="font-size: 11px;">available</div>
                </div>
              </div>
            </div>
          `
          : nothing
      }
    </section>
  `;
}
