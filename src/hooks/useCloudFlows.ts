import { useQuery } from "@tanstack/react-query";
import { dataverseGet, isPowerAppsRuntime } from "@/hooks/useDataverseCatalog";
import { mapDataverseRecords } from "@/lib/summaryConfiguration";
import type { StudioConfiguration } from "@/hooks/useStudioRuntime";

export type CloudFlowState = "Draft" | "Activated" | "Suspended" | "Unknown";

export interface CloudFlow {
  id: string;
  name: string;
  description: string;
  state: CloudFlowState;
  createdOn: string | null;
  modifiedOn: string | null;
}

export interface RelatedCloudFlow extends CloudFlow {
  /** True when the configuration record points at this flow. */
  linked: boolean;
}

const fallbackFlows: CloudFlow[] = [
  {
    id: "mock-flow-1",
    name: "PPS - Account operations summary - 0329",
    description: "Generated from the Account operations summary recipe.",
    state: "Activated",
    createdOn: new Date(Date.now() - 86_400_000 * 3).toISOString(),
    modifiedOn: new Date(Date.now() - 86_400_000).toISOString(),
  },
  {
    id: "mock-flow-2",
    name: "PPS - Account operations summary - 0851",
    description: "Earlier generation of the same recipe.",
    state: "Activated",
    createdOn: new Date(Date.now() - 86_400_000 * 7).toISOString(),
    modifiedOn: new Date(Date.now() - 86_400_000 * 6).toISOString(),
  },
];

function stateFromCode(value: unknown): CloudFlowState {
  switch (Number(value)) {
    case 0:
      return "Draft";
    case 1:
      return "Activated";
    case 2:
      return "Suspended";
    default:
      return "Unknown";
  }
}

/** Cloud flows (workflow records of category 5) in the current environment. */
export function useCloudFlows() {
  return useQuery({
    queryKey: ["cloud-flows"],
    staleTime: 30_000,
    queryFn: async (): Promise<CloudFlow[]> => {
      if (!isPowerAppsRuntime()) return fallbackFlows;
      const result = await dataverseGet(
        "StudioListCloudFlows",
        {},
        "Dataverse could not list the cloud flows in this environment.",
      );
      return mapDataverseRecords(result).map((raw) => ({
        id: String(raw.workflowid ?? ""),
        name: String(raw.name ?? ""),
        description: String(raw.description ?? ""),
        state: stateFromCode(raw.statecode),
        createdOn: raw.createdon ? String(raw.createdon) : null,
        modifiedOn: raw.modifiedon ? String(raw.modifiedon) : null,
      }));
    },
  });
}

/**
 * Flows generated for a configuration: the one linked through
 * `csp_processid`/`csp_flowid`, plus any flow whose name carries the recipe
 * name or technical flow name (earlier generations the backend left behind).
 */
export function relatedCloudFlows(
  flows: CloudFlow[],
  configuration: Pick<StudioConfiguration, "id" | "name" | "flowId" | "flowName">,
): RelatedCloudFlow[] {
  const linkedId = configuration.flowId.toLowerCase();
  const recipeName = configuration.name.trim().toLowerCase();
  const technicalName = configuration.flowName.trim().toLowerCase();
  const linked = flows.find((flow) => linkedId && flow.id.toLowerCase() === linkedId);
  // Generated flows share the linked flow's base name and differ only by a
  // numeric suffix ("PPS - Recipe - 0329"), which survives recipe renames.
  const linkedBase = linked ? linked.name.replace(/\s*-\s*\d+\s*$/, "").trim().toLowerCase() : "";
  return flows
    .filter((flow) => {
      const name = flow.name.toLowerCase();
      if (linkedId && flow.id.toLowerCase() === linkedId) return true;
      if (linkedBase && name.startsWith(linkedBase)) return true;
      if (recipeName && name.includes(recipeName)) return true;
      if (technicalName && name.includes(technicalName)) return true;
      return false;
    })
    .map((flow) => ({ ...flow, linked: Boolean(linkedId) && flow.id.toLowerCase() === linkedId }))
    .sort((a, b) => {
      if (a.linked !== b.linked) return a.linked ? -1 : 1;
      return String(b.modifiedOn ?? "").localeCompare(String(a.modifiedOn ?? ""));
    });
}

/** Unique name of the Power Platform solution that holds the studio tables and generated flows. */
export const STUDIO_SOLUTION_UNIQUE_NAME = "BizzSummit2026";

export interface StudioSolution {
  id: string;
  uniqueName: string;
  friendlyName: string;
  version: string;
}

export function useStudioSolution() {
  return useQuery({
    queryKey: ["studio-solution", STUDIO_SOLUTION_UNIQUE_NAME],
    staleTime: 10 * 60 * 1000,
    queryFn: async (): Promise<StudioSolution | null> => {
      if (!isPowerAppsRuntime()) return null;
      const result = await dataverseGet(
        "StudioGetSolution",
        { uniqueName: STUDIO_SOLUTION_UNIQUE_NAME },
        "Dataverse could not read the studio solution.",
      );
      const raw = mapDataverseRecords(result)[0];
      if (!raw) return null;
      return {
        id: String(raw.solutionid ?? ""),
        uniqueName: String(raw.uniquename ?? STUDIO_SOLUTION_UNIQUE_NAME),
        friendlyName: String(raw.friendlyname ?? STUDIO_SOLUTION_UNIQUE_NAME),
        version: String(raw.version ?? ""),
      };
    },
  });
}

export function solutionUrl(environmentId: string, solutionId: string) {
  if (!environmentId) return "";
  return solutionId
    ? `https://make.powerapps.com/environments/${environmentId}/solutions/${solutionId}`
    : `https://make.powerapps.com/environments/${environmentId}/solutions`;
}

export function cloudFlowUrl(environmentId: string, flowId: string) {
  if (!environmentId || !flowId) return "";
  return `https://make.powerautomate.com/environments/${environmentId}/flows/${flowId}/details`;
}

export function cloudFlowRunUrl(environmentId: string, flowId: string, runId: string) {
  if (!environmentId || !flowId || !runId) return "";
  return `https://make.powerautomate.com/environments/${environmentId}/flows/${flowId}/runs/${runId}`;
}
