import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";

export interface StudioPrompt {
  id: string;
  name: string;
  key: string;
  model: string;
  version: string;
  content: string;
}

export interface StudioConfiguration {
  id: string;
  name: string;
  description: string;
  entity: string;
  sourceFields: string[];
  triggerColumns: string[];
  relationships: Array<Record<string, unknown>>;
  inputMappings: Record<string, string>;
  outputEntity: string;
  outputField: string;
  model: string;
  version: string;
  flowName: string;
  status: number;
  mode: number;
  promptId: string;
  promptName: string;
  enabled: boolean;
  lastRun: string | null;
}

export interface StudioRun {
  id: string;
  name: string;
  accountName: string;
  model: string;
  timestamp: string | null;
  latencyMs: number;
  tokens: number;
  status: number;
}

export interface StudioAccount {
  id: string;
  name: string;
  summary: string;
  generatedOn: string | null;
}

export interface StudioRuntime {
  live: boolean;
  configurations: StudioConfiguration[];
  prompts: StudioPrompt[];
  runs: StudioRun[];
  accounts: StudioAccount[];
  cacheCount: number;
}

const fallback: StudioRuntime = {
  live: false,
  configurations: [{
    id: "mock-config",
    name: "Account operations summary",
    description: "Primary Summary Studio demo.",
    entity: "account",
    sourceFields: ["name", "industrycode", "revenue", "description", "primarycontactid", "modifiedon"],
    triggerColumns: ["name", "revenue", "description", "primarycontactid"],
    relationships: [{ entity: "activitypointer", windowDays: 30, maxRecords: 12 }],
    inputMappings: { account_context: "compiled.primaryAndRelated", generated_at: "utcNow" },
    outputEntity: "account",
    outputField: "csp_aisummary",
    model: "gpt-4.1-mini",
    version: "3.0",
    flowName: "csp_SUM_AccountOperations_v3",
    status: 100000001,
    mode: 100000000,
    promptId: "mock-prompt",
    promptName: "Account executive summary",
    enabled: true,
    lastRun: new Date().toISOString(),
  }],
  prompts: [{ id: "mock-prompt", name: "Account executive summary", key: "account-operations", model: "gpt-4.1-mini", version: "3.0", content: "Generate an executive summary using {{account_context}}." }],
  runs: [
    { id: "mock-run-1", name: "RUN-9284", accountName: "Contoso Retail", model: "gpt-4.1-mini", timestamp: new Date().toISOString(), latencyMs: 2850, tokens: 760, status: 100000000 },
    { id: "mock-run-2", name: "RUN-9283", accountName: "Northwind Health", model: "gpt-4.1-mini", timestamp: new Date(Date.now() - 3600000).toISOString(), latencyMs: 3160, tokens: 814, status: 100000000 },
    { id: "mock-run-3", name: "RUN-9282", accountName: "Fabrikam Energy", model: "gpt-4.1-mini", timestamp: new Date(Date.now() - 7200000).toISOString(), latencyMs: 3470, tokens: 868, status: 100000000 },
  ],
  accounts: [
    { id: "mock-account-1", name: "Contoso Retail", summary: "Eight stores affected; engineering is investigating policy replication.", generatedOn: new Date().toISOString() },
    { id: "mock-account-2", name: "Northwind Health", summary: "Upcoming renewal is awaiting SLA evidence.", generatedOn: new Date(Date.now() - 3600000).toISOString() },
    { id: "mock-account-3", name: "Fabrikam Energy", summary: "Regional expansion depends on the adoption plan.", generatedOn: new Date(Date.now() - 7200000).toISOString() },
  ],
  cacheCount: 3,
};

const runtimeKey = ["summary-studio-runtime"] as const;

function parseJson<T>(value: unknown, fallbackValue: T): T {
  if (typeof value !== "string" || !value) return fallbackValue;
  try { return JSON.parse(value) as T; } catch { return fallbackValue; }
}

function isPowerAppsRuntime() {
  if (typeof window === "undefined") return false;
  return window.parent !== window || window.location.hostname.endsWith("powerapps.com");
}

export function useStudioRuntime() {
  return useQuery({
    queryKey: runtimeKey,
    queryFn: async (): Promise<StudioRuntime> => {
      if (!isPowerAppsRuntime()) return fallback;
      try {
        const [accountsModule, promptsModule, configsModule, cachesModule, usagesModule] = await Promise.all([
          import("@/generated/services/AccountsService"),
          import("@/generated/services/Csp_aipromptsService"),
          import("@/generated/services/Csp_aisummaryconfigsService"),
          import("@/generated/services/Csp_aisummarycachesService"),
          import("@/generated/services/Csp_aiusagesService"),
        ]);
        const { AccountsService } = accountsModule;
        const { Csp_aipromptsService } = promptsModule;
        const { Csp_aisummaryconfigsService } = configsModule;
        const { Csp_aisummarycachesService } = cachesModule;
        const { Csp_aiusagesService } = usagesModule;
        const [configResult, promptResult, runResult, accountResult, cacheResult] = await Promise.all([
          Csp_aisummaryconfigsService.getAll({ filter: "statecode eq 0", top: 100 }),
          Csp_aipromptsService.getAll({ filter: "statecode eq 0", top: 100 }),
          Csp_aiusagesService.getAll({ filter: "statecode eq 0", top: 50 }),
          AccountsService.getAll({ filter: "accountnumber ne null and startswith(accountnumber,'SUM-DEMO-')", top: 24 }),
          Csp_aisummarycachesService.getAll({ filter: "statecode eq 0", top: 100 }),
        ]);
        const prompts = (promptResult.data ?? []).map((item) => {
          const raw = item as unknown as Record<string, unknown>;
          const key = String(raw.csp_promptkey ?? "");
          const isSummitDemo = key === "account-operations";
          return { id: String(raw.csp_aipromptid ?? ""), name: isSummitDemo ? "Account executive summary" : String(raw.csp_name ?? ""), key, model: String(raw.csp_model ?? ""), version: String(raw.csp_version ?? ""), content: isSummitDemo ? "Generate an executive summary using {{account_context}}." : String(raw.csp_content ?? "") };
        });
        const configurations = (configResult.data ?? []).map((item) => {
          const raw = item as unknown as Record<string, unknown>;
          const isSummitDemo = String(raw.csp_flowname ?? "") === "csp_SUM_AccountOperations_v3";
          return {
            id: String(raw.csp_aisummaryconfigid ?? ""), name: isSummitDemo ? "Account operations summary" : String(raw.csp_name ?? ""), description: isSummitDemo ? "Primary Summary Studio demo." : String(raw.csp_description ?? ""), entity: String(raw.csp_targetentity ?? "account"),
            sourceFields: parseJson<string[]>(raw.csp_sourcefields, []), triggerColumns: parseJson<string[]>(raw.csp_triggercolumns, []), relationships: parseJson<Array<Record<string, unknown>>>(raw.csp_relationships, []), inputMappings: parseJson<Record<string, string>>(raw.csp_inputmappings, {}),
            outputEntity: String(raw.csp_outputentity ?? raw.csp_targetentity ?? "account"), outputField: String(raw.csp_outputfield ?? "csp_aisummary"), model: String(raw.csp_model ?? "gpt-4.1-mini"), version: String(raw.csp_version ?? "1.0"), flowName: String(raw.csp_flowname ?? ""),
            status: Number(raw.csp_status ?? 100000000), mode: Number(raw.csp_mode ?? 100000000), promptId: String(raw._csp_prompt_value ?? ""), promptName: isSummitDemo ? "Account executive summary" : String(raw.csp_promptname ?? ""), enabled: Boolean(raw.csp_enabled), lastRun: raw.csp_lastrun ? String(raw.csp_lastrun) : null,
          };
        });
        const accounts = (accountResult.data ?? []).map((item) => {
          const raw = item as unknown as Record<string, unknown>;
          return { id: String(raw.accountid ?? ""), name: String(raw.name ?? ""), summary: String(raw.csp_aisummary ?? ""), generatedOn: raw.csp_aisummarygeneratedon ? String(raw.csp_aisummarygeneratedon) : null };
        });
        const accountNames = new Map(accounts.map((account) => [account.id.toLowerCase(), account.name]));
        const runs = (runResult.data ?? []).map((item) => {
          const raw = item as unknown as Record<string, unknown>;
          const targetId = String(raw.csp_targetrecordid ?? "").toLowerCase();
          return { id: String(raw.csp_aiusageid ?? ""), name: String(raw.csp_name ?? ""), accountName: accountNames.get(targetId) ?? String(raw.csp_accountname ?? "Dataverse record"), model: String(raw.csp_model ?? ""), timestamp: raw.csp_timestamp ? String(raw.csp_timestamp) : null, latencyMs: Number(raw.csp_latencyms ?? 0), tokens: Number(raw.csp_tokensinput ?? 0) + Number(raw.csp_tokensoutput ?? 0), status: Number(raw.csp_status ?? 100000000) };
        }).sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
        return { live: true, configurations, prompts, runs, accounts, cacheCount: (cacheResult.data ?? []).length };
      } catch (error) {
        console.error("Summary Studio could not connect to Dataverse; using local fallback.", error);
        return fallback;
      }
    },
    staleTime: 15_000,
  });
}

export function useUpdateStudioConfiguration() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, changes }: { id: string; changes: Record<string, unknown> }) => {
      if (id === "mock-config" || !isPowerAppsRuntime()) return;
      const { Csp_aisummaryconfigsService } = await import("@/generated/services/Csp_aisummaryconfigsService");
      await Csp_aisummaryconfigsService.update(id, changes as never);
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}
