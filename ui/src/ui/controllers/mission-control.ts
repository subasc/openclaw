import type { GatewayBrowserClient } from "../gateway.ts";
import type { CostUsageSummary, MissionControlBridge, MissionControlSnapshot } from "../types.ts";

export type MissionControlState = {
  client: GatewayBrowserClient | null;
  connected: boolean;
  missionControlLoading: boolean;
  missionControlError: string | null;
  missionControlSnapshot: MissionControlSnapshot | null;
  missionControlBridges: MissionControlBridge[] | null;
  missionControlCostSummary: CostUsageSummary | null;
  missionControlLastRefresh: number | null;
};

export async function loadMissionControl(state: MissionControlState) {
  if (!state.client || !state.connected) {
    return;
  }
  if (state.missionControlLoading) {
    return;
  }
  state.missionControlLoading = true;
  state.missionControlError = null;
  try {
    const [snapshot, bridgeResult, costResult] = await Promise.all([
      state.client.request("mission-control.snapshot", {}),
      state.client.request("unified-inbox.status", {}).catch(() => null),
      state.client.request("usage.cost", { days: 30 }).catch(() => null),
    ]);
    state.missionControlSnapshot = snapshot as MissionControlSnapshot;
    if (bridgeResult && typeof bridgeResult === "object" && "bots" in bridgeResult) {
      state.missionControlBridges = (bridgeResult as { bots: MissionControlBridge[] }).bots;
    }
    if (costResult && typeof costResult === "object") {
      state.missionControlCostSummary = costResult as CostUsageSummary;
    }
    state.missionControlLastRefresh = Date.now();
  } catch (err) {
    state.missionControlError = String(err);
  } finally {
    state.missionControlLoading = false;
  }
}
