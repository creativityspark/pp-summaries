export type SummaryConfigurationDraft = {
  name: string;
  entity: string;
  sourceFields: string[];
  outputEntity: string;
  outputField: string;
  model: string;
  mode: number;
  promptId: string;
  promptName: string;
  promptKey: string;
  promptContent: string;
  flowName: string;
  relationships: Array<Record<string, unknown>>;
  inputMappings: Record<string, string>;
  triggerColumns: string[];
  triggerType: string;
  preserveHistory: boolean;
  saveMetadata: boolean;
  queryMode: number;
  maxRecords: number;
  fetchXml: string;
  relatedFetchXml: string;
};

const primaryIdByEntity: Record<string, string> = {
  account: "accountid",
  incident: "incidentid",
  opportunity: "opportunityid",
};

const entitySetByEntity: Record<string, string> = {
  account: "accounts",
  incident: "incidents",
  opportunity: "opportunities",
};

export function getTargetEntityIdentity(entity: string) {
  return {
    entitySet: entitySetByEntity[entity] ?? `${entity}s`,
    idField: primaryIdByEntity[entity] ?? `${entity}id`,
  };
}

export function buildRecordSelectionFetchXml(entity: string, maxRecords: number) {
  const primaryId = primaryIdByEntity[entity] ?? `${entity}id`;
  return `<fetch top="${maxRecords}"><entity name="${entity}"><attribute name="${primaryId}" /></entity></fetch>`;
}

export function buildContextFetchXml(
  entity: string,
  fields: string[],
  relationships: Array<Record<string, unknown>>,
) {
  const primaryId = primaryIdByEntity[entity] ?? `${entity}id`;
  const attributes = Array.from(new Set([primaryId, ...fields]))
    .map((field) => `<attribute name="${field}" />`)
    .join("");
  const links = relationships.map((relationship, index) => {
    const relatedEntity = String(relationship.entity ?? "activitypointer");
    const windowDays = Number(relationship.windowDays ?? 30);
    const from = relatedEntity === "activitypointer" ? "regardingobjectid" : `${entity}id`;
    const configuredQuery = String(relationship.fetchXml ?? "");
    const configuredFilter = configuredQuery.match(/<filter\b[\s\S]*?<\/filter>/i)?.[0];
    const configuredOrder = configuredQuery.match(/<order\b[^>]*\/?\s*>/i)?.[0] ?? "";
    const relatedFilter = configuredFilter ?? `<filter><condition attribute="createdon" operator="last-x-days" value="${windowDays}" /></filter>`;
    return `<link-entity name="${relatedEntity}" from="${from}" to="${primaryId}" alias="related${index + 1}" link-type="outer"><attribute name="activityid" /><attribute name="subject" /><attribute name="description" />${relatedFilter}${configuredOrder}</link-entity>`;
  }).join("");
  return `<fetch><entity name="${entity}">${attributes}<filter><condition attribute="${primaryId}" operator="eq" value="{{recordId}}" /></filter>${links}</entity></fetch>`;
}

export function compileRecipe(draft: SummaryConfigurationDraft) {
  return {
    version: "3.0",
    source: {
      entity: draft.entity,
      fields: draft.sourceFields,
      relationships: draft.relationships,
      recordSelection: {
        queryMode: draft.queryMode,
        maxRecords: draft.maxRecords,
        fetchXml: draft.fetchXml,
      },
      contextFetchXml: draft.relatedFetchXml,
    },
    prompt: {
      id: draft.promptId,
      key: draft.promptKey,
      model: draft.model,
      inputs: draft.inputMappings,
    },
    output: {
      entity: draft.outputEntity,
      field: draft.outputField,
      cache: draft.saveMetadata,
      preserveHistory: draft.preserveHistory,
    },
    trigger: {
      type: draft.triggerType,
      columns: draft.triggerColumns,
    },
  };
}

export function buildConfigurationPayload(draft: SummaryConfigurationDraft): Record<string, unknown> {
  const target = getTargetEntityIdentity(draft.entity);
  const payload: Record<string, unknown> = {
    csp_name: draft.name,
    csp_targetentity: draft.entity,
    csp_targetentityset: target.entitySet,
    csp_targetentityidfield: target.idField,
    csp_sourcefields: JSON.stringify(draft.sourceFields),
    csp_relationships: JSON.stringify(draft.relationships),
    csp_inputmappings: JSON.stringify(draft.inputMappings),
    csp_outputentity: draft.outputEntity,
    csp_outputfield: draft.outputField,
    csp_model: draft.model,
    csp_mode: draft.mode,
    csp_triggercolumns: JSON.stringify(draft.triggerColumns),
    csp_fetchxml: draft.fetchXml,
    csp_relatedfetchxml: draft.relatedFetchXml,
    csp_querymode: draft.queryMode,
    csp_maxrecords: draft.maxRecords,
    csp_version: "3.0",
    csp_enabled: true,
    csp_status: 100000001,
    csp_configurationjson: JSON.stringify(compileRecipe(draft)),
  };

  if (draft.flowName) payload.csp_flowname = draft.flowName;
  if (draft.promptId) payload["csp_Prompt@odata.bind"] = `/csp_aiprompts(${draft.promptId})`;

  return payload;
}

export function buildDraftConfigurationPayload(draft: SummaryConfigurationDraft): Record<string, unknown> {
  return {
    ...buildConfigurationPayload(draft),
    csp_enabled: false,
    csp_status: 100000000,
    statecode: 0,
  };
}

export function validateSummaryConfiguration(draft: SummaryConfigurationDraft): string[] {
  const errors: string[] = [];
  if (!draft.name.trim()) errors.push("Give this summary a name.");
  if (!draft.entity.trim()) errors.push("Select a source table.");
  if (draft.sourceFields.length === 0) errors.push("Select at least one source field.");
  if (!draft.promptId.trim()) errors.push("Select an AI Prompt.");
  if (!draft.fetchXml.trim().startsWith("<fetch")) errors.push("Provide valid record-selection FetchXML.");
  if (!draft.relatedFetchXml.trim().startsWith("<fetch")) errors.push("Provide valid context FetchXML.");
  else if (draft.mode === 100000000 && !draft.relatedFetchXml.includes("{{recordId}}")) errors.push("The context FetchXML must include {{recordId}}.");
  if (!draft.outputEntity.trim()) errors.push("Select a destination table.");
  if (!draft.outputField.trim()) errors.push("Select a destination column.");
  if (draft.triggerType === "dataverse.update" && draft.triggerColumns.length === 0) errors.push("Select at least one trigger column.");
  if (draft.triggerType === "dataverse.update" && draft.triggerColumns.includes(draft.outputField)) errors.push("Remove the destination column from the trigger to prevent a loop.");
  return errors;
}

export function buildPublishSignal(): Record<string, unknown> {
  return { statuscode: 787000001 };
}
