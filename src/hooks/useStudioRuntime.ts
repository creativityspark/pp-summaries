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
  entitySetName: string;
  entityIdField: string;
  sourceFields: string[];
  triggerColumns: string[];
  relationships: Array<Record<string, unknown>>;
  inputMappings: Record<string, string>;
  outputEntity: string;
  outputEntitySetName: string;
  outputField: string;
  model: string;
  version: string;
  /** Id of the generated cloud flow (csp_flowid or the csp_processid lookup). */
  flowId: string;
  /** Technical flow name stored by the studio (csp_flowname). */
  flowName: string;
  /** Display name of the generated cloud flow, when the backend linked one. */
  flowDisplayName: string;
  status: number;
  mode: number;
  promptId: string;
  promptName: string;
  promptKey: string;
  enabled: boolean;
  lastRun: string | null;
  lastRunStatus: number | null;
  lastRunError: string;
  queryMode: number;
  systemViewId?: string;
  userViewId?: string;
  maxRecords: number;
  fetchXml: string;
  relatedFetchXml: string;
}

export interface StudioRun {
  id: string;
  name: string;
  configurationId: string;
  configurationName: string;
  accountName: string;
  targetEntity: string;
  targetRecordId: string;
  promptName: string;
  model: string;
  timestamp: string | null;
  latencyMs: number;
  tokens: number;
  status: number;
  flowRunId: string;
  error: string;
}

export interface StudioAccount {
  id: string;
  name: string;
  summary: string;
  generatedOn: string | null;
}

export interface StudioUser {
  name: string;
  email: string;
}

export interface StudioRuntime {
  live: boolean;
  configurations: StudioConfiguration[];
  prompts: StudioPrompt[];
  runs: StudioRun[];
  accounts: StudioAccount[];
  cacheCount: number;
  user: StudioUser | null;
  environmentId: string;
  orgUrl: string;
}

/** Local dataset used when the app runs outside the Power Apps host. */
const fallback: StudioRuntime = {
  live: false,
  user: null,
  environmentId: "",
  orgUrl: "",
  configurations: [
    {
      id: "mock-config",
      name: "Account operations summary",
      description: "Primary Summary Studio demo.",
      entity: "account",
      entitySetName: "accounts",
      entityIdField: "accountid",
      sourceFields: [
        "name",
        "industrycode",
        "revenue",
        "description",
        "primarycontactid",
        "modifiedon",
      ],
      triggerColumns: ["name", "revenue", "description", "primarycontactid"],
      relationships: [{ entity: "activitypointer", windowDays: 30, maxRecords: 12 }],
      inputMappings: {
        account_context: "compiled.primaryAndRelated",
        generated_at: "utcNow",
      },
      outputEntity: "account",
      outputEntitySetName: "accounts",
      outputField: "csp_aisummary",
      model: "gpt-4.1-mini",
      version: "3.0",
      flowId: "mock-flow-1",
      flowName: "csp_SUM_AccountOperations_v3",
      flowDisplayName: "PPS - Account operations summary - 0329",
      status: 100000001,
      mode: 100000000,
      promptId: "mock-prompt",
      promptName: "Account executive summary",
      promptKey: "account-operations",
      enabled: true,
      lastRun: new Date().toISOString(),
      lastRunStatus: 100000000,
      lastRunError: "",
      queryMode: 100000002,
      maxRecords: 5000,
      fetchXml:
        '<fetch top="5000"><entity name="account"><attribute name="accountid" /></entity></fetch>',
      relatedFetchXml:
        '<fetch><entity name="account"><attribute name="accountid" /><attribute name="name" /><filter><condition attribute="accountid" operator="eq" value="{{recordId}}" /></filter></entity></fetch>',
    },
  ],
  prompts: [
    {
      id: "mock-prompt",
      name: "Account executive summary",
      key: "account-operations",
      model: "gpt-4.1-mini",
      version: "3.0",
      content: "Generate an executive summary using {{account_context}}.",
    },
  ],
  runs: [
    {
      id: "mock-run-1",
      name: "RUN-9284",
      configurationId: "mock-config",
      configurationName: "Account operations summary",
      accountName: "Contoso Retail",
      targetEntity: "account",
      targetRecordId: "mock-account-1",
      promptName: "Account executive summary",
      model: "gpt-4.1-mini",
      flowRunId: "08584400000000000000000000000CU01",
      error: "",
      timestamp: new Date().toISOString(),
      latencyMs: 2850,
      tokens: 760,
      status: 100000000,
    },
    {
      id: "mock-run-2",
      name: "RUN-9283",
      configurationId: "mock-config",
      configurationName: "Account operations summary",
      accountName: "Northwind Health",
      targetEntity: "account",
      targetRecordId: "mock-account-2",
      promptName: "Account executive summary",
      model: "gpt-4.1-mini",
      flowRunId: "08584400000000000000000000000CU02",
      error: "",
      timestamp: new Date(Date.now() - 3600000).toISOString(),
      latencyMs: 3160,
      tokens: 814,
      status: 100000000,
    },
    {
      id: "mock-run-3",
      name: "RUN-9282",
      configurationId: "mock-config",
      configurationName: "Account operations summary",
      accountName: "Fabrikam Energy",
      targetEntity: "account",
      targetRecordId: "mock-account-3",
      promptName: "Account executive summary",
      model: "gpt-4.1-mini",
      flowRunId: "08584400000000000000000000000CU03",
      error: "",
      timestamp: new Date(Date.now() - 7200000).toISOString(),
      latencyMs: 3470,
      tokens: 868,
      status: 100000000,
    },
  ],
  accounts: [
    {
      id: "mock-account-1",
      name: "Contoso Retail",
      summary: "Eight stores affected; engineering is investigating policy replication.",
      generatedOn: new Date().toISOString(),
    },
    {
      id: "mock-account-2",
      name: "Northwind Health",
      summary: "Upcoming renewal is awaiting SLA evidence.",
      generatedOn: new Date(Date.now() - 3600000).toISOString(),
    },
    {
      id: "mock-account-3",
      name: "Fabrikam Energy",
      summary: "Regional expansion depends on the adoption plan.",
      generatedOn: new Date(Date.now() - 7200000).toISOString(),
    },
  ],
  cacheCount: 3,
};

const runtimeKey = ["summary-studio-runtime"] as const;

type Raw = Record<string, unknown>;

function parseJson<T>(value: unknown, fallbackValue: T): T {
  if (typeof value !== "string" || !value) return fallbackValue;
  try {
    return JSON.parse(value) as T;
  } catch {
    return fallbackValue;
  }
}

function text(value: unknown, fallbackValue = ""): string {
  if (value === null || value === undefined) return fallbackValue;
  const result = String(value);
  return result || fallbackValue;
}

function isPowerAppsRuntime() {
  if (typeof window === "undefined") return false;
  return (
    window.parent !== window ||
    window.location.hostname.endsWith("powerapps.com")
  );
}

async function readHostContext(): Promise<{
  user: StudioUser | null;
  environmentId: string;
  orgUrl: string;
}> {
  try {
    const { getContext } = await import("@microsoft/power-apps/app");
    const context = await getContext();
    const name = context.user?.fullName?.trim() ?? "";
    const email = context.user?.userPrincipalName ?? "";
    return {
      user: name || email ? { name: name || email, email } : null,
      environmentId: context.app?.environmentId ?? "",
      orgUrl: context.app?.dataverseOrgUrl ?? "",
    };
  } catch {
    return { user: null, environmentId: "", orgUrl: "" };
  }
}

async function loadDataverseRuntime(): Promise<StudioRuntime> {
  const [
    { AccountsService },
    { Csp_aipromptsService },
    { Csp_aisummaryconfigsService },
    { Csp_aisummarycachesService },
    { Csp_aiusagesService },
  ] = await Promise.all([
    import("@/generated/services/AccountsService"),
    import("@/generated/services/Csp_aipromptsService"),
    import("@/generated/services/Csp_aisummaryconfigsService"),
    import("@/generated/services/Csp_aisummarycachesService"),
    import("@/generated/services/Csp_aiusagesService"),
  ]);

  const [configResult, promptResult, runResult, accountResult, cacheResult, host] =
    await Promise.all([
      Csp_aisummaryconfigsService.getAll({ filter: "statecode eq 0", top: 100 }),
      Csp_aipromptsService.getAll({ filter: "statecode eq 0", top: 100 }),
      Csp_aiusagesService.getAll({ filter: "statecode eq 0", top: 50 }),
      AccountsService.getAll({
        filter: "accountnumber ne null and startswith(accountnumber,'SUM-DEMO-')",
        top: 24,
      }),
      Csp_aisummarycachesService.getAll({ filter: "statecode eq 0", top: 100 }),
      readHostContext(),
    ]);

  const prompts: StudioPrompt[] = (promptResult.data ?? []).map((item) => {
    const raw = item as unknown as Raw;
    return {
      id: text(raw.csp_aipromptid),
      name: text(raw.csp_name, "Untitled prompt"),
      key: text(raw.csp_promptkey),
      model: text(raw.csp_model),
      version: text(raw.csp_version),
      content: text(raw.csp_content),
    };
  });
  const promptsById = new Map(prompts.map((prompt) => [prompt.id.toLowerCase(), prompt]));

  const configurations: StudioConfiguration[] = (configResult.data ?? []).map((item) => {
    const raw = item as unknown as Raw;
    const compiled = parseJson<{ output?: { entitySetName?: string } }>(
      raw.csp_configurationjson,
      {},
    );
    const promptId = text(raw._csp_prompt_value);
    const prompt = promptsById.get(promptId.toLowerCase());
    const entity = text(raw.csp_targetentity, "account");
    return {
      id: text(raw.csp_aisummaryconfigid),
      name: text(raw.csp_name, "Untitled configuration"),
      description: text(raw.csp_description),
      entity,
      entitySetName: text(raw.csp_targetentityset, "accounts"),
      entityIdField: text(raw.csp_targetentityidfield, "accountid"),
      sourceFields: parseJson<string[]>(raw.csp_sourcefields, []),
      triggerColumns: parseJson<string[]>(raw.csp_triggercolumns, []),
      relationships: parseJson<Array<Raw>>(raw.csp_relationships, []),
      inputMappings: parseJson<Record<string, string>>(raw.csp_inputmappings, {}),
      outputEntity: text(raw.csp_outputentity, entity),
      outputEntitySetName: text(compiled.output?.entitySetName),
      outputField: text(raw.csp_outputfield, "csp_aisummary"),
      model: text(raw.csp_model, prompt?.model || "gpt-4.1-mini"),
      version: text(raw.csp_version, "1.0"),
      flowId: text(raw.csp_flowid) || text(raw._csp_processid_value),
      flowName: text(raw.csp_flowname),
      flowDisplayName: text(
        raw["_csp_processid_value@OData.Community.Display.V1.FormattedValue"] ??
          raw.csp_processidname,
      ),
      status: Number(raw.csp_status ?? 100000000),
      mode: Number(raw.csp_mode ?? 100000000),
      promptId,
      promptName: prompt?.name ?? "",
      promptKey: prompt?.key ?? "",
      enabled: Boolean(raw.csp_enabled),
      lastRun: raw.csp_lastrun ? text(raw.csp_lastrun) : null,
      lastRunStatus:
        raw.csp_lastrunstatus === null || raw.csp_lastrunstatus === undefined
          ? null
          : Number(raw.csp_lastrunstatus),
      lastRunError: text(raw.csp_lastrunerror),
      queryMode: Number(raw.csp_querymode ?? 100000002),
      maxRecords: Number(raw.csp_maxrecords ?? 5000),
      fetchXml: text(raw.csp_fetchxml),
      relatedFetchXml: text(raw.csp_relatedfetchxml),
      systemViewId: raw.csp_systemviewid ? text(raw.csp_systemviewid) : undefined,
      userViewId: raw.csp_userviewid ? text(raw.csp_userviewid) : undefined,
    };
  });
  const configurationsById = new Map(
    configurations.map((config) => [config.id.toLowerCase(), config]),
  );

  const accounts: StudioAccount[] = (accountResult.data ?? []).map((item) => {
    const raw = item as unknown as Raw;
    return {
      id: text(raw.accountid),
      name: text(raw.name),
      summary: text(raw.csp_aisummary),
      generatedOn: raw.csp_aisummarygeneratedon ? text(raw.csp_aisummarygeneratedon) : null,
    };
  });
  const accountNames = new Map(accounts.map((account) => [account.id.toLowerCase(), account.name]));

  const runs: StudioRun[] = (runResult.data ?? [])
    .map((item) => {
      const raw = item as unknown as Raw;
      const targetId = text(raw.csp_targetrecordid).toLowerCase();
      const accountId = text(raw._csp_account_value).toLowerCase();
      const configurationId = text(raw._csp_summaryconfig_value);
      const configuration = configurationsById.get(configurationId.toLowerCase());
      return {
        id: text(raw.csp_aiusageid),
        name: text(raw.csp_name),
        configurationId,
        configurationName:
          configuration?.name ?? text(raw.csp_summaryconfigname, "Summary"),
        accountName:
          accountNames.get(targetId) ??
          accountNames.get(accountId) ??
          text(raw.csp_accountname, "Dataverse record"),
        targetEntity: text(raw.csp_targetentity),
        targetRecordId: text(raw.csp_targetrecordid) || text(raw._csp_account_value),
        promptName: text(raw.csp_promptlookupname) || text(raw.csp_promptname),
        model: text(raw.csp_model),
        timestamp: raw.csp_timestamp ? text(raw.csp_timestamp) : null,
        latencyMs: Number(raw.csp_latencyms ?? 0),
        tokens: Number(raw.csp_tokensinput ?? 0) + Number(raw.csp_tokensoutput ?? 0),
        status: Number(raw.csp_status ?? 100000000),
        flowRunId: text(raw.csp_flowrunid),
        error: text(raw.csp_error),
      };
    })
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));

  return {
    live: true,
    configurations,
    prompts,
    runs,
    accounts,
    cacheCount: (cacheResult.data ?? []).length,
    user: host.user,
    environmentId: host.environmentId,
    orgUrl: host.orgUrl,
  };
}

export function useStudioRuntime() {
  return useQuery({
    queryKey: runtimeKey,
    queryFn: async (): Promise<StudioRuntime> => {
      if (!isPowerAppsRuntime()) return fallback;
      try {
        return await loadDataverseRuntime();
      } catch (error) {
        console.error(
          "Summary Studio could not connect to Dataverse; using local fallback.",
          error,
        );
        return fallback;
      }
    },
    staleTime: 15_000,
  });
}

function ensureSuccess(result: unknown, fallbackMessage: string) {
  const outcome = result as { success?: boolean; error?: { message?: string } } | null;
  if (outcome && outcome.success === false) {
    throw new Error(outcome.error?.message?.trim() || fallbackMessage);
  }
}

export interface StudioPromptInput {
  name: string;
  key: string;
  model: string;
  version: string;
  content: string;
}

function toPromptFields(input: Partial<StudioPromptInput>): Raw {
  const fields: Raw = {};
  if (input.name !== undefined) fields.csp_name = input.name;
  if (input.key !== undefined) fields.csp_promptkey = input.key;
  if (input.model !== undefined) fields.csp_model = input.model;
  if (input.version !== undefined) fields.csp_version = input.version;
  if (input.content !== undefined) fields.csp_content = input.content;
  return fields;
}

export function useSaveStudioPrompt() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, changes }: { id: string; changes: Partial<StudioPromptInput> }) => {
      if (id.startsWith("mock-") || !isPowerAppsRuntime()) return;
      const { Csp_aipromptsService } = await import(
        "@/generated/services/Csp_aipromptsService"
      );
      const result = await Csp_aipromptsService.update(id, toPromptFields(changes) as never);
      ensureSuccess(result, "Dataverse could not update the AI Prompt.");
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}

export function useCreateStudioPrompt() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (input: StudioPromptInput) => {
      if (!isPowerAppsRuntime()) return "mock-new-prompt";
      const { Csp_aipromptsService } = await import(
        "@/generated/services/Csp_aipromptsService"
      );
      const result = await Csp_aipromptsService.create({
        ...toPromptFields(input),
        csp_enabled: true,
      } as never);
      ensureSuccess(result, "Dataverse could not create the AI Prompt.");
      const id = text((result.data as unknown as Raw | undefined)?.csp_aipromptid);
      if (!id) throw new Error("Dataverse did not return the new AI Prompt identifier.");
      return id;
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}

export function useUpdateStudioConfiguration() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, changes }: { id: string; changes: Record<string, unknown> }) => {
      if (id === "mock-config" || !isPowerAppsRuntime()) return;
      const { Csp_aisummaryconfigsService } = await import(
        "@/generated/services/Csp_aisummaryconfigsService"
      );
      const result = await Csp_aisummaryconfigsService.update(id, changes as never);
      ensureSuccess(result, "Dataverse could not update the configuration.");
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}

export function useCreateStudioConfiguration() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (changes: Record<string, unknown>) => {
      if (!isPowerAppsRuntime()) return "mock-new-config";
      const { Csp_aisummaryconfigsService } = await import(
        "@/generated/services/Csp_aisummaryconfigsService"
      );
      const result = await Csp_aisummaryconfigsService.create(changes as never);
      ensureSuccess(result, "Dataverse could not create the configuration.");
      const data = result.data as unknown as Raw | undefined;
      const id = text(data?.csp_aisummaryconfigid);
      if (!id) {
        throw new Error("Dataverse did not return the new summary configuration identifier.");
      }
      return id;
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}

export function useDeactivateStudioConfiguration() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      if (id.startsWith("mock-") || !isPowerAppsRuntime()) return;
      const { Csp_aisummaryconfigsService } = await import(
        "@/generated/services/Csp_aisummaryconfigsService"
      );
      const result = await Csp_aisummaryconfigsService.update(id, {
        statecode: 1,
        statuscode: 2,
        csp_enabled: false,
      } as never);
      ensureSuccess(result, "Dataverse could not deactivate the configuration.");
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}

export function useDeleteStudioConfiguration() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      if (id.startsWith("mock-") || !isPowerAppsRuntime()) return;
      const { Csp_aisummaryconfigsService } = await import(
        "@/generated/services/Csp_aisummaryconfigsService"
      );
      await Csp_aisummaryconfigsService.delete(id);
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}

export function useUpdateStudioPrompt() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, content }: { id: string; content: string }) => {
      if (id === "mock-prompt" || !isPowerAppsRuntime()) return;
      const { Csp_aipromptsService } = await import(
        "@/generated/services/Csp_aipromptsService"
      );
      const result = await Csp_aipromptsService.update(id, { csp_content: content } as never);
      ensureSuccess(result, "Dataverse could not update the AI Prompt.");
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: runtimeKey }),
  });
}
