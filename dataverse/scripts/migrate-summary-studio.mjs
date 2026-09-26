import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const SOLUTION = "BizzSummit2026";

const obsoleteTables = [
  "csp_userfeedback",
  "csp_contactlogentry",
  "csp_nextbestaction",
  "csp_accounthealth",
  "csp_weeklyrecap",
];

const additions = {
  account: [
    { type: "Memo", schema: "csp_AISummary", label: "AI Summary", maxLength: 32000 },
    { type: "DateTime", schema: "csp_AISummaryGeneratedOn", label: "AI Summary Generated On" },
  ],
  csp_aisummaryconfig: [
    { type: "Memo", schema: "csp_SourceFields", label: "Source Fields", maxLength: 32000 },
    { type: "Memo", schema: "csp_Relationships", label: "Relationships", maxLength: 32000 },
    { type: "Memo", schema: "csp_InputMappings", label: "Input Mappings", maxLength: 32000 },
    { type: "Memo", schema: "csp_TriggerColumns", label: "Trigger Columns", maxLength: 4000 },
    { type: "String", schema: "csp_OutputEntity", label: "Output Entity", maxLength: 100 },
    { type: "String", schema: "csp_OutputField", label: "Output Field", maxLength: 100 },
    { type: "String", schema: "csp_Model", label: "Model", maxLength: 100 },
    { type: "String", schema: "csp_Version", label: "Version", maxLength: 50 },
    { type: "String", schema: "csp_FlowId", label: "Cloud Flow Id", maxLength: 100 },
    { type: "String", schema: "csp_FlowName", label: "Cloud Flow Name", maxLength: 200 },
    { type: "Memo", schema: "csp_ConfigurationJson", label: "Compiled Configuration", maxLength: 32000 },
    { type: "Picklist", schema: "csp_Status", label: "Configuration Status", options: ["Draft", "Published", "Paused", "Error"] },
  ],
  csp_aisummarycache: [
    { type: "String", schema: "csp_OutputEntity", label: "Output Entity", maxLength: 100 },
    { type: "String", schema: "csp_OutputField", label: "Output Field", maxLength: 100 },
    { type: "Integer", schema: "csp_DurationMs", label: "Duration (ms)", min: 0, max: 3600000 },
    { type: "Memo", schema: "csp_Error", label: "Error", maxLength: 4000 },
    { type: "Picklist", schema: "csp_Status", label: "Generation Status", options: ["Success", "Failed", "Blocked"] },
  ],
  csp_aiusage: [
    { type: "String", schema: "csp_FlowRunId", label: "Cloud Flow Run Id", maxLength: 200 },
    { type: "Memo", schema: "csp_Error", label: "Error", maxLength: 4000 },
    { type: "Picklist", schema: "csp_Status", label: "Execution Status", options: ["Success", "Failed", "Blocked"] },
  ],
};

const pca = new PublicClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
});
const token = (await pca.acquireTokenByDeviceCode({
  scopes: [`${ENV_URL}/.default`],
  deviceCodeCallback: (response) => {
    console.log("\nDEVICE CODE AUTHENTICATION");
    console.log(response.message);
    console.log("");
  },
})).accessToken;

async function request(method, path, body, { solution = false, returnRepresentation = false } = {}) {
  const response = await fetch(`${ENV_URL}/api/data/v9.2${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      ...(solution ? { "MSCRM.SolutionUniqueName": SOLUTION } : {}),
      ...(returnRepresentation ? { Prefer: "return=representation" } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!response.ok) throw new Error(`${method} ${path} -> ${response.status}\n${await response.text()}`);
  if (response.status === 204) {
    const entityId = response.headers.get("odata-entityid") ?? response.headers.get("OData-EntityId");
    return entityId ? { id: entityId.match(/\(([^)]+)\)$/)?.[1] } : null;
  }
  return response.json();
}

const label = (text) => ({ LocalizedLabels: [{ Label: text, LanguageCode: 1033 }], UserLocalizedLabel: { Label: text, LanguageCode: 1033 } });

function attributeBody(definition) {
  const common = {
    SchemaName: definition.schema,
    DisplayName: label(definition.label),
    Description: label(`Summary Studio: ${definition.label}.`),
    RequiredLevel: { Value: "None" },
  };
  if (definition.type === "String") return { "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata", AttributeType: "String", AttributeTypeName: { Value: "StringType" }, MaxLength: definition.maxLength, FormatName: { Value: "Text" }, ...common };
  if (definition.type === "Memo") return { "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata", AttributeType: "Memo", AttributeTypeName: { Value: "MemoType" }, MaxLength: definition.maxLength, Format: "TextArea", ...common };
  if (definition.type === "Integer") return { "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata", AttributeType: "Integer", AttributeTypeName: { Value: "IntegerType" }, Format: "None", MinValue: definition.min, MaxValue: definition.max, ...common };
  if (definition.type === "DateTime") return { "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata", AttributeType: "DateTime", AttributeTypeName: { Value: "DateTimeType" }, Format: "DateAndTime", DateTimeBehavior: { Value: "UserLocal" }, ...common };
  if (definition.type === "Picklist") return { "@odata.type": "Microsoft.Dynamics.CRM.PicklistAttributeMetadata", AttributeType: "Picklist", AttributeTypeName: { Value: "PicklistType" }, OptionSet: { "@odata.type": "Microsoft.Dynamics.CRM.OptionSetMetadata", OptionSetType: "Picklist", IsGlobal: false, Options: definition.options.map((name, index) => ({ Value: 100000000 + index, Label: label(name) })) }, ...common };
  throw new Error(`Unsupported attribute type ${definition.type}`);
}

async function exists(path) {
  try { await request("GET", path); return true; } catch (error) { if (error.message.includes("-> 404")) return false; throw error; }
}

async function deleteRows(entitySet, primaryKey, filter = "") {
  let path = `/${entitySet}?$select=${primaryKey}${filter ? `&$filter=${encodeURIComponent(filter)}` : ""}`;
  let deleted = 0;
  while (path) {
    const page = await request("GET", path);
    for (const row of page.value) {
      await request("DELETE", `/${entitySet}(${row[primaryKey]})`);
      deleted += 1;
    }
    path = page["@odata.nextLink"]?.replace(`${ENV_URL}/api/data/v9.2`, "") ?? "";
  }
  return deleted;
}

async function createRecord(entitySet, body) {
  const result = await request("POST", `/${entitySet}`, body, { returnRepresentation: true });
  const key = Object.keys(result).find((name) => name.endsWith("id") && typeof result[name] === "string");
  return { id: result[key], row: result };
}

console.log("1/6 Removing obsolete scenario tables...");
for (const logicalName of obsoleteTables) {
  if (!(await exists(`/EntityDefinitions(LogicalName='${logicalName}')?$select=MetadataId`))) {
    console.log(`  [skip] ${logicalName}`);
    continue;
  }
  const metadata = await request("GET", `/EntityDefinitions(LogicalName='${logicalName}')?$select=MetadataId`);
  await request("DELETE", `/EntityDefinitions(${metadata.MetadataId})`);
  console.log(`  [deleted] ${logicalName}`);
}

console.log("2/6 Removing old seeded accounts and contacts...");
if (await exists("/EntityDefinitions(LogicalName='contact')/Attributes(LogicalName='csp_isdemo')?$select=MetadataId")) {
  console.log(`  contacts: ${await deleteRows("contacts", "contactid", "csp_isdemo eq true")}`);
}
if (await exists("/EntityDefinitions(LogicalName='account')/Attributes(LogicalName='csp_isdemo')?$select=MetadataId")) {
  console.log(`  accounts: ${await deleteRows("accounts", "accountid", "csp_isdemo eq true")}`);
}
for (const entity of ["contact", "account"]) {
  if (await exists(`/EntityDefinitions(LogicalName='${entity}')/Attributes(LogicalName='csp_isdemo')?$select=MetadataId`)) {
    await request("DELETE", `/EntityDefinitions(LogicalName='${entity}')/Attributes(LogicalName='csp_isdemo')`);
    console.log(`  [deleted attribute] ${entity}.csp_isdemo`);
  }
}

console.log("3/6 Clearing reusable Summary Studio tables...");
for (const table of [
  ["csp_aiusages", "csp_aiusageid"],
  ["csp_aisummarycaches", "csp_aisummarycacheid"],
  ["csp_aisummaryconfigs", "csp_aisummaryconfigid"],
  ["csp_aiprompts", "csp_aipromptid"],
]) console.log(`  ${table[0]}: ${await deleteRows(table[0], table[1])}`);

console.log("4/6 Adding the Summary Studio schema...");
for (const [entity, definitions] of Object.entries(additions)) {
  for (const definition of definitions) {
    const logicalName = definition.schema.toLowerCase();
    if (await exists(`/EntityDefinitions(LogicalName='${entity}')/Attributes(LogicalName='${logicalName}')?$select=MetadataId`)) {
      console.log(`  [skip] ${entity}.${logicalName}`);
      continue;
    }
    await request("POST", `/EntityDefinitions(LogicalName='${entity}')/Attributes?SolutionUniqueName=${SOLUTION}`, attributeBody(definition), { solution: true });
    console.log(`  [created] ${entity}.${logicalName}`);
  }
}

console.log("5/6 Seeding a real, non-aviation demo...");
const prompt = await createRecord("csp_aiprompts", {
  csp_name: "Account executive summary",
  csp_promptkey: "account-operations",
  csp_description: "Summarizes the operational and commercial status of an account using Dataverse data only.",
  csp_systemmessage: "You are an account analyst. Do not invent information and explicitly flag unknown data.",
  csp_content: "Generate an executive summary in English with the sections: current situation, recent changes, risks, opportunities, and next action. Use only {{account_context}}. Generated at: {{generated_at}}.",
  csp_expectedvariables: JSON.stringify(["account_context", "generated_at"]),
  csp_outputformat: 100000000,
  csp_model: "gpt-4.1-mini",
  csp_deploymentname: "gpt-4.1-mini",
  csp_version: "3.0",
  csp_enabled: true,
  csp_temperature: 0.2,
  csp_maxtokens: 700,
});

const sourceFields = ["name", "industrycode", "revenue", "description", "primarycontactid", "modifiedon"];
const relationships = [{ entity: "activitypointer", relationship: "regardingobjectid", windowDays: 30, maxRecords: 12 }];
const inputMappings = { account_context: "compiled.primaryAndRelated", generated_at: "utcNow" };
const fetchXml = "<fetch top='24'><entity name='account'><attribute name='accountid'/><attribute name='name'/><attribute name='industrycode'/><attribute name='revenue'/><attribute name='description'/><attribute name='primarycontactid'/><attribute name='modifiedon'/><order attribute='modifiedon' descending='true'/></entity></fetch>";
const relatedFetchXml = "<fetch top='12'><entity name='activitypointer'><attribute name='subject'/><attribute name='description'/><attribute name='activitytypecode'/><attribute name='modifiedon'/><filter><condition attribute='regardingobjectid' operator='eq' value='{{recordId}}'/><condition attribute='modifiedon' operator='last-x-days' value='30'/></filter><order attribute='modifiedon' descending='true'/></entity></fetch>";
const compiled = { version: "3.0", source: { entity: "account", fields: sourceFields, fetchXml, relationships }, prompt: { id: prompt.id, key: "account-operations", model: "gpt-4.1-mini", inputs: inputMappings }, output: { entity: "account", field: "csp_aisummary", cache: true }, trigger: { type: "dataverse.update", columns: ["name", "revenue", "description", "primarycontactid"] } };
const config = await createRecord("csp_aisummaryconfigs", {
  csp_name: "Account operations summary",
  csp_description: "Primary Summary Studio demo for the Irish Power Platform Summit 2026.",
  csp_cadence: 100000000,
  csp_mode: 100000000,
  "csp_Prompt@odata.bind": `/csp_aiprompts(${prompt.id})`,
  csp_targetentity: "account",
  csp_querymode: 100000002,
  csp_fetchxml: fetchXml,
  csp_relatedfetchxml: relatedFetchXml,
  csp_maxrecords: 24,
  csp_language: "es-ES",
  csp_outputdestination: 100000000,
  csp_outputcontext: JSON.stringify({ entity: "account", field: "csp_aisummary", cache: true }),
  csp_enabled: true,
  csp_sourcefields: JSON.stringify(sourceFields),
  csp_relationships: JSON.stringify(relationships),
  csp_inputmappings: JSON.stringify(inputMappings),
  csp_triggercolumns: JSON.stringify(["name", "revenue", "description", "primarycontactid"]),
  csp_outputentity: "account",
  csp_outputfield: "csp_aisummary",
  csp_model: "gpt-4.1-mini",
  csp_version: "3.0",
  csp_flowname: "csp_SUM_AccountOperations_v3",
  csp_configurationjson: JSON.stringify(compiled),
  csp_status: 100000001,
});

const accounts = [];
for (const row of [
  { name: "Contoso Retail", accountnumber: "SUM-DEMO-001", revenue: 4800000, description: "Expanding retail chain. Current priority: stabilize store authentication and prepare the opening of two new locations." },
  { name: "Northwind Health", accountnumber: "SUM-DEMO-002", revenue: 7200000, description: "Healthcare provider with an upcoming renewal. Requests greater traceability of response times and compliance." },
  { name: "Fabrikam Energy", accountnumber: "SUM-DEMO-003", revenue: 9600000, description: "Energy group evaluating a regional expansion. There is an opportunity conditional on closing the adoption plan." },
]) accounts.push(await createRecord("accounts", row));

const taskData = [
  [0, "Review authentication impact", "Eight stores report intermittent failures; existing sessions remain stable."],
  [0, "Confirm plan with engineering", "Policy replication is the leading hypothesis, still unconfirmed."],
  [1, "Prepare renewal committee", "The customer requests SLA evidence and named owners before the renewal."],
  [2, "Close adoption plan", "The expansion depends on agreeing milestones, owners, and dates."],
];
for (const [accountIndex, subject, description] of taskData) {
  await createRecord("tasks", { subject, description, "regardingobjectid_account_task@odata.bind": `/accounts(${accounts[accountIndex].id})` });
}

const now = new Date();
for (let index = 0; index < accounts.length; index += 1) {
  const account = accounts[index];
  const summary = index === 0
    ? "Contoso Retail has eight stores affected by intermittent authentication failures. Existing sessions remain stable. Engineering is investigating policy replication; the diagnosis is not yet confirmed. Next action: validate telemetry and send the customer an update."
    : index === 1
      ? "Northwind Health is approaching a significant renewal. The customer is requesting SLA evidence, named owners, and a follow-up plan before the committee. Next action: consolidate metrics and confirm owners."
      : "Fabrikam Energy is evaluating expanding the service to a new region. The opportunity depends on closing an adoption plan with milestones and owners. Next action: send the revised plan for approval.";
  await request("PATCH", `/accounts(${account.id})`, { csp_aisummary: summary, csp_aisummarygeneratedon: new Date(now.getTime() - index * 3600000).toISOString() });
  const cache = await createRecord("csp_aisummarycaches", {
    csp_name: `${["Contoso Retail", "Northwind Health", "Fabrikam Energy"][index]} · Operational summary`,
    "csp_Account@odata.bind": `/accounts(${account.id})`,
    "csp_SummaryConfig@odata.bind": `/csp_aisummaryconfigs(${config.id})`,
    "csp_Prompt@odata.bind": `/csp_aiprompts(${prompt.id})`,
    csp_targetentity: "account",
    csp_targetrecordid: account.id,
    csp_summarytype: 100000000,
    csp_content: summary,
    csp_generatedon: new Date(now.getTime() - index * 3600000).toISOString(),
    csp_tokensused: 760 + index * 54,
    csp_promptversion: "3.0",
    csp_outputentity: "account",
    csp_outputfield: "csp_aisummary",
    csp_durationms: 2850 + index * 310,
    csp_status: 100000000,
  });
  await createRecord("csp_aiusages", {
    csp_name: `RUN-${9284 - index} · ${["Contoso Retail", "Northwind Health", "Fabrikam Energy"][index]}`,
    "csp_PromptLookup@odata.bind": `/csp_aiprompts(${prompt.id})`,
    "csp_SummaryConfig@odata.bind": `/csp_aisummaryconfigs(${config.id})`,
    "csp_Account@odata.bind": `/accounts(${account.id})`,
    csp_promptname: "Account executive summary",
    csp_targetentity: "account",
    csp_targetrecordid: account.id,
    csp_tokensinput: 620 + index * 35,
    csp_tokensoutput: 140 + index * 19,
    csp_estimatedcost: 0.0009 + index * 0.0001,
    csp_latencyms: 2850 + index * 310,
    csp_model: "gpt-4.1-mini",
    csp_timestamp: new Date(now.getTime() - index * 3600000).toISOString(),
    csp_flowrunid: `mock-flow-run-${9284 - index}`,
    csp_status: 100000000,
  });
  void cache;
}

console.log("6/6 Publishing metadata...");
await request("POST", "/PublishAllXml", {});
console.log("\nSummary Studio migration completed successfully.");
console.log(`Prompt: ${prompt.id}`);
console.log(`Configuration: ${config.id}`);
console.log(`Demo accounts: ${accounts.map((account) => account.id).join(", ")}`);
