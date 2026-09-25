export type SummaryConfigurationDraft = {
  name: string;
  entity: string;
  entitySetName?: string;
  entityIdField?: string;
  sourceFields: string[];
  outputEntity: string;
  outputEntitySetName?: string;
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
  systemViewId?: string;
  userViewId?: string;
  maxRecords: number;
  fetchXml: string;
  relatedFetchXml: string;
};

export interface DataverseTableMetadata {
  logicalName: string;
  entitySetName: string;
  primaryIdAttribute: string;
  displayName: string;
}

export interface DataverseColumnMetadata {
  logicalName: string;
  displayName: string;
  type: string;
  readable: boolean;
  writable: boolean;
}

export interface DataverseRelationshipMetadata {
  schemaName: string;
  entity: string;
  fromAttribute: string;
  toAttribute: string;
}

function localizedLabel(value: unknown): string {
  if (typeof value === "string") return value;
  if (!value || typeof value !== "object") return "";
  const label = value as Record<string, unknown>;
  const userLabel = label.UserLocalizedLabel ?? label.userLocalizedLabel;
  if (userLabel && typeof userLabel === "object") {
    const text = (userLabel as Record<string, unknown>).Label ?? (userLabel as Record<string, unknown>).label;
    if (typeof text === "string") return text;
  }
  const direct = label.Label ?? label.label;
  return typeof direct === "string" ? direct : "";
}

function responseValue(value: unknown): unknown[] {
  if (!value || typeof value !== "object") return [];
  const root = value as Record<string, unknown>;
  const data = root.data && typeof root.data === "object" ? root.data as Record<string, unknown> : root;
  const items = data.value ?? data.Value;
  return Array.isArray(items) ? items : [];
}

export function mapDataverseTables(value: unknown): DataverseTableMetadata[] {
  return responseValue(value).map((item) => {
    const record = item as Record<string, unknown>;
    const raw = record.dynamicProperties && typeof record.dynamicProperties === "object"
      ? record.dynamicProperties as Record<string, unknown>
      : record;
    const logicalName = String(raw.LogicalName ?? raw.logicalName ?? "");
    const entitySetName = String(raw.EntitySetName ?? raw.entitySetName ?? "");
    const primaryIdAttribute = String(raw.PrimaryIdAttribute ?? raw.primaryIdAttribute ?? `${logicalName}id`);
    const displayName = localizedLabel(raw.DisplayCollectionName ?? raw.displayCollectionName) || logicalName;
    return { logicalName, entitySetName, primaryIdAttribute, displayName };
  }).filter((table) => table.logicalName && table.entitySetName)
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export function mapDataverseTableDefinition(value: unknown, fallback: DataverseTableMetadata): DataverseTableMetadata {
  const root = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const data = root.data && typeof root.data === "object" ? root.data as Record<string, unknown> : root;
  const records = mapDataverseRecords(value);
  const raw = records[0] ?? data;
  return {
    logicalName: String(raw.LogicalName ?? raw.logicalName ?? fallback.logicalName),
    entitySetName: String(raw.EntitySetName ?? raw.entitySetName ?? fallback.entitySetName),
    primaryIdAttribute: String(raw.PrimaryIdAttribute ?? raw.primaryIdAttribute ?? fallback.primaryIdAttribute),
    displayName: localizedLabel(raw.DisplayCollectionName ?? raw.displayCollectionName) || fallback.displayName,
  };
}

export function mapDataverseColumns(value: unknown): DataverseColumnMetadata[] {
  if (!value || typeof value !== "object") return [];
  const root = value as Record<string, unknown>;
  const data = root.data && typeof root.data === "object" ? root.data as Record<string, unknown> : root;
  const attributes = data.Attributes ?? data.attributes;
  if (!Array.isArray(attributes)) return [];
  return attributes.map((item) => {
    const raw = item as Record<string, unknown>;
    const logicalName = String(raw.LogicalName ?? raw.logicalName ?? "");
    return {
      logicalName,
      displayName: localizedLabel(raw.DisplayName ?? raw.displayName) || logicalName,
      type: String(raw.AttributeType ?? raw.attributeType ?? raw.type ?? "Unknown"),
      readable: raw.IsValidForRead === undefined ? raw.isValidForRead !== false : raw.IsValidForRead !== false,
      writable: raw.IsValidForUpdate === true || raw.isValidForUpdate === true,
    };
  }).filter((column) => column.logicalName)
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export function mapDataverseRecords(value: unknown): Array<Record<string, unknown>> {
  return responseValue(value).map((item) => {
    const raw = item as Record<string, unknown>;
    const dynamic = raw.dynamicProperties;
    return dynamic && typeof dynamic === "object" ? dynamic as Record<string, unknown> : raw;
  });
}

export function mapDataverseRelationships(value: unknown, sourceEntity: string): DataverseRelationshipMetadata[] {
  if (!value || typeof value !== "object") return [];
  const root = value as Record<string, unknown>;
  const data = root.data && typeof root.data === "object" ? root.data as Record<string, unknown> : root;
  const oneToMany = data.OneToManyRelationships ?? data.oneToManyRelationships;
  const manyToOne = data.ManyToOneRelationships ?? data.manyToOneRelationships;
  const relationships: DataverseRelationshipMetadata[] = [];

  if (Array.isArray(oneToMany)) {
    for (const item of oneToMany) {
      const raw = item as Record<string, unknown>;
      const referencedEntity = String(raw.ReferencedEntity ?? raw.referencedEntity ?? "");
      if (referencedEntity !== sourceEntity) continue;
      relationships.push({
        schemaName: String(raw.SchemaName ?? raw.schemaName ?? ""),
        entity: String(raw.ReferencingEntity ?? raw.referencingEntity ?? ""),
        fromAttribute: String(raw.ReferencingAttribute ?? raw.referencingAttribute ?? ""),
        toAttribute: String(raw.ReferencedAttribute ?? raw.referencedAttribute ?? ""),
      });
    }
  }

  if (Array.isArray(manyToOne)) {
    for (const item of manyToOne) {
      const raw = item as Record<string, unknown>;
      const referencingEntity = String(raw.ReferencingEntity ?? raw.referencingEntity ?? "");
      if (referencingEntity !== sourceEntity) continue;
      relationships.push({
        schemaName: String(raw.SchemaName ?? raw.schemaName ?? ""),
        entity: String(raw.ReferencedEntity ?? raw.referencedEntity ?? ""),
        fromAttribute: String(raw.ReferencedAttribute ?? raw.referencedAttribute ?? ""),
        toAttribute: String(raw.ReferencingAttribute ?? raw.referencingAttribute ?? ""),
      });
    }
  }

  return relationships.filter((relationship) => relationship.entity && relationship.fromAttribute && relationship.toAttribute);
}

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

export function buildRecordSelectionFetchXml(entity: string, maxRecords: number, primaryIdAttribute?: string) {
  const primaryId = primaryIdAttribute || primaryIdByEntity[entity] || `${entity}id`;
  return `<fetch top="${maxRecords}"><entity name="${entity}"><attribute name="${primaryId}" /></entity></fetch>`;
}

export function buildContextFetchXml(
  entity: string,
  fields: string[],
  relationships: Array<Record<string, unknown>>,
  primaryIdAttribute?: string,
) {
  const primaryId = primaryIdAttribute || primaryIdByEntity[entity] || `${entity}id`;
  const attributes = Array.from(new Set([primaryId, ...fields]))
    .map((field) => `<attribute name="${field}" />`)
    .join("");
  const links = relationships.map((relationship, index) => {
    const relatedEntity = String(relationship.entity ?? "activitypointer");
    const windowDays = Number(relationship.windowDays ?? 30);
    const from = String(relationship.fromAttribute ?? (relatedEntity === "activitypointer" ? "regardingobjectid" : `${entity}id`));
    const to = String(relationship.toAttribute ?? primaryId);
    const relatedFields = Array.isArray(relationship.fields)
      ? relationship.fields.filter((field): field is string => typeof field === "string")
      : ["activityid", "subject", "description"];
    const relatedAttributes = relatedFields.map((field) => `<attribute name="${field}" />`).join("");
    const configuredQuery = String(relationship.fetchXml ?? "");
    const configuredFilter = configuredQuery.match(/<filter\b[\s\S]*?<\/filter>/i)?.[0];
    const configuredOrder = configuredQuery.match(/<order\b[^>]*\/?\s*>/i)?.[0] ?? "";
    const relatedFilter = configuredFilter ?? `<filter><condition attribute="createdon" operator="last-x-days" value="${windowDays}" /></filter>`;
    return `<link-entity name="${relatedEntity}" from="${from}" to="${to}" alias="related${index + 1}" link-type="outer">${relatedAttributes}${relatedFilter}${configuredOrder}</link-entity>`;
  }).join("");
  return `<fetch><entity name="${entity}">${attributes}<filter><condition attribute="${primaryId}" operator="eq" value="{{recordId}}" /></filter>${links}</entity></fetch>`;
}

export function compileRecipe(draft: SummaryConfigurationDraft) {
  const target = getTargetEntityIdentity(draft.entity);
  return {
    version: "3.0",
    source: {
      entity: draft.entity,
      entitySetName: draft.entitySetName || target.entitySet,
      primaryIdAttribute: draft.entityIdField || target.idField,
      fields: draft.sourceFields,
      relationships: draft.relationships,
      recordSelection: {
        queryMode: draft.queryMode,
        maxRecords: draft.maxRecords,
        fetchXml: draft.fetchXml,
        ...(draft.queryMode === 100000000 && draft.systemViewId ? { viewId: draft.systemViewId, viewType: "system" } : {}),
        ...(draft.queryMode === 100000001 && draft.userViewId ? { viewId: draft.userViewId, viewType: "personal" } : {}),
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
      entitySetName: draft.outputEntitySetName || getTargetEntityIdentity(draft.outputEntity).entitySet,
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
    csp_targetentityset: draft.entitySetName || target.entitySet,
    csp_targetentityidfield: draft.entityIdField || target.idField,
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

  if (draft.systemViewId) payload.csp_systemviewid = draft.systemViewId;
  if (draft.userViewId) payload.csp_userviewid = draft.userViewId;

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
