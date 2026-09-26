import { getClient } from "@microsoft/power-apps/data";
import { useMutation, useQuery } from "@tanstack/react-query";
import { dataSourcesInfo } from "../../.power/schemas/appschemas/dataSourcesInfo";
import {
  mapDataverseColumns,
  mapDataverseRelationships,
  mapDataverseRecords,
  mapDataverseTableDefinition,
  mapDataverseTables,
  type DataverseColumnMetadata,
  type DataverseRelationshipMetadata,
  type DataverseTableMetadata,
} from "@/lib/summaryConfiguration";

const client = getClient(dataSourcesInfo);

/**
 * Native Dataverse Web API operations for the studio catalog.
 *
 * The generic connector operation (`commondataserviceforapps.ListRecords`)
 * cannot be used here: the app's connector reference has no dataset, so the
 * connector runtime rejects the call with "Invalid organization URL 'null'".
 * The SDK's Dataverse executor, which the generated services use, resolves
 * the organization from the app's `default.cds` database reference instead.
 * It only runs operations declared on a Dataverse data source, so we register
 * these GET operations on the `accounts` data source at module load. The
 * SDK keeps a single shared reference to `dataSourcesInfo`, so every client
 * created from it sees the additional operations. Path placeholders are
 * substituted with `encodeURIComponent`, which is why query values live in
 * the path template.
 */
const NATIVE_DATA_SOURCE = "accounts";
const ENTITY_SELECT = "LogicalName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,DisplayCollectionName";
const ENTITY_EXPAND =
  "Attributes($select=LogicalName,DisplayName,AttributeType,IsValidForRead,IsValidForUpdate)," +
  "OneToManyRelationships($select=SchemaName,ReferencedEntity,ReferencedAttribute,ReferencingEntity,ReferencingAttribute)," +
  "ManyToOneRelationships($select=SchemaName,ReferencedEntity,ReferencedAttribute,ReferencingEntity,ReferencingAttribute)";

type NativeApiDefinition = {
  path: string;
  method: string;
  parameters: Array<{ name: string; in: string; required: boolean; type: string; format?: string }>;
};

const nativeApis: Record<string, NativeApiDefinition> = {
  StudioRunFetchXml: {
    path: "/api/data/v9.0/{entitySetName}?fetchXml={fetchXml}&$top={top}",
    method: "GET",
    parameters: [
      { name: "entitySetName", in: "path", required: true, type: "string", format: "dataverse-entity-set-name" },
      { name: "fetchXml", in: "path", required: true, type: "string" },
      { name: "top", in: "path", required: true, type: "string" },
    ],
  },
  StudioListEntityDefinitions: {
    path: `/api/data/v9.0/EntityDefinitions?$select=${encodeURIComponent(ENTITY_SELECT)}`,
    method: "GET",
    parameters: [],
  },
  StudioListCloudFlows: {
    path:
      "/api/data/v9.0/workflows?$select=workflowid,name,description,statecode,statuscode,category,type,createdon,modifiedon" +
      "&$filter=category%20eq%205%20and%20type%20eq%201&$orderby=modifiedon%20desc&$top=200",
    method: "GET",
    parameters: [],
  },
  StudioGetSolution: {
    path: "/api/data/v9.0/solutions?$select=solutionid,uniquename,friendlyname,version&$filter=uniquename%20eq%20'{uniqueName}'",
    method: "GET",
    parameters: [{ name: "uniqueName", in: "path", required: true, type: "string" }],
  },
  StudioGetEntityDefinition: {
    path: `/api/data/v9.0/EntityDefinitions(LogicalName='{logicalName}')?$select=${encodeURIComponent(ENTITY_SELECT)}&$expand=${encodeURIComponent(ENTITY_EXPAND)}`,
    method: "GET",
    parameters: [{ name: "logicalName", in: "path", required: true, type: "string" }],
  },
};

const nativeDataSource = (dataSourcesInfo as Record<string, { apis: Record<string, unknown> }>)[NATIVE_DATA_SOURCE];
if (nativeDataSource) {
  for (const [name, definition] of Object.entries(nativeApis)) {
    if (!(name in nativeDataSource.apis)) nativeDataSource.apis[name] = definition;
  }
}

export async function dataverseGet(operationName: keyof typeof nativeApis, body: Record<string, string>, failureMessage: string) {
  const result = await client.executeAsync<unknown, unknown>({
    dataverseRequest: {
      action: "customapi",
      parameters: { operationName, tableName: NATIVE_DATA_SOURCE, body },
    },
  });
  unwrapResult(result, failureMessage);
  return result;
}

const fallbackTables: DataverseTableMetadata[] = [
  { logicalName: "account", entitySetName: "accounts", primaryIdAttribute: "accountid", displayName: "Accounts", primaryNameAttribute: "name" },
  { logicalName: "contact", entitySetName: "contacts", primaryIdAttribute: "contactid", displayName: "Contacts", primaryNameAttribute: "fullname" },
  { logicalName: "incident", entitySetName: "incidents", primaryIdAttribute: "incidentid", displayName: "Cases", primaryNameAttribute: "title" },
  { logicalName: "opportunity", entitySetName: "opportunities", primaryIdAttribute: "opportunityid", displayName: "Opportunities", primaryNameAttribute: "name" },
].sort((a, b) => a.displayName.localeCompare(b.displayName));

const fallbackColumns: Record<string, DataverseColumnMetadata[]> = {
  account: [
    { logicalName: "name", displayName: "Account Name", type: "String", readable: true, writable: true },
    { logicalName: "accountnumber", displayName: "Account Number", type: "String", readable: true, writable: true },
    { logicalName: "description", displayName: "Description", type: "Memo", readable: true, writable: true },
    { logicalName: "revenue", displayName: "Annual Revenue", type: "Money", readable: true, writable: true },
    { logicalName: "industrycode", displayName: "Industry", type: "Picklist", readable: true, writable: true },
    { logicalName: "primarycontactid", displayName: "Primary Contact", type: "Lookup", readable: true, writable: true },
    { logicalName: "modifiedon", displayName: "Modified On", type: "DateTime", readable: true, writable: false },
    { logicalName: "csp_aisummary", displayName: "AI Summary", type: "Memo", readable: true, writable: true },
  ],
  contact: [
    { logicalName: "fullname", displayName: "Full Name", type: "String", readable: true, writable: false },
    { logicalName: "firstname", displayName: "First Name", type: "String", readable: true, writable: true },
    { logicalName: "lastname", displayName: "Last Name", type: "String", readable: true, writable: true },
    { logicalName: "emailaddress1", displayName: "Email", type: "String", readable: true, writable: true },
    { logicalName: "description", displayName: "Description", type: "Memo", readable: true, writable: true },
  ],
  incident: [
    { logicalName: "title", displayName: "Case Title", type: "String", readable: true, writable: true },
    { logicalName: "ticketnumber", displayName: "Case Number", type: "String", readable: true, writable: false },
    { logicalName: "description", displayName: "Description", type: "Memo", readable: true, writable: true },
  ],
  opportunity: [
    { logicalName: "name", displayName: "Topic", type: "String", readable: true, writable: true },
    { logicalName: "estimatedvalue", displayName: "Estimated Revenue", type: "Money", readable: true, writable: true },
    { logicalName: "description", displayName: "Description", type: "Memo", readable: true, writable: true },
  ],
};

const fallbackRelationships: Record<string, DataverseRelationshipMetadata[]> = {
  account: [
    { schemaName: "account_contacts", entity: "contact", fromAttribute: "parentcustomerid", toAttribute: "accountid" },
    { schemaName: "account_opportunities", entity: "opportunity", fromAttribute: "parentaccountid", toAttribute: "accountid" },
    { schemaName: "Account_ActivityPointers", entity: "activitypointer", fromAttribute: "regardingobjectid", toAttribute: "accountid" },
  ],
};

export function isPowerAppsRuntime() {
  if (typeof window === "undefined") return false;
  return window.parent !== window || window.location.hostname.endsWith("powerapps.com");
}

/** Runs FetchXML against Dataverse through the native Web API path. */
export async function runFetchXml(
  entitySetName: string,
  fetchXml: string,
  limit: number,
  failureMessage = "Dataverse could not execute this query.",
) {
  const safeLimit = Math.max(1, Math.floor(limit));
  const result = await dataverseGet(
    "StudioRunFetchXml",
    { entitySetName, fetchXml: prepareFetchXmlForPreview(fetchXml), top: String(safeLimit) },
    failureMessage,
  );
  return mapDataverseRecords(result).slice(0, safeLimit);
}

async function loadEntityMetadata(logicalName: string) {
  const result = await dataverseGet(
    "StudioGetEntityDefinition",
    { logicalName: logicalName.replaceAll("'", "''") },
    "Dataverse metadata request failed.",
  );
  const data = (result as { data?: unknown }).data;
  if (data && typeof data === "object" && !Array.isArray(data) && "LogicalName" in data) {
    return data as Record<string, unknown>;
  }
  return mapDataverseRecords(result)[0] ?? {};
}

export function useDataverseTables() {
  return useQuery({
    queryKey: ["dataverse-table-catalog"],
    staleTime: 10 * 60 * 1000,
    queryFn: async () => {
      if (!isPowerAppsRuntime()) return fallbackTables;
      try {
        const result = await dataverseGet(
          "StudioListEntityDefinitions",
          {},
          "Dataverse table catalog request failed.",
        );
        const tables = mapDataverseTables(result);
        return tables.length ? tables : fallbackTables;
      } catch {
        return fallbackTables;
      }
    },
  });
}

export function useDataverseTableMetadata(table?: DataverseTableMetadata) {
  return useQuery({
    queryKey: ["dataverse-table-metadata", table?.logicalName],
    enabled: Boolean(table),
    staleTime: 10 * 60 * 1000,
    queryFn: async () => {
      if (!table || !isPowerAppsRuntime()) return table;
      try {
        const metadata = await loadEntityMetadata(table.logicalName);
        return mapDataverseTableDefinition(metadata, table);
      } catch {
        return table;
      }
    },
  });
}

export function useDataverseColumns(logicalName: string, entitySetName: string) {
  return useQuery({
    queryKey: ["dataverse-columns", logicalName, entitySetName],
    enabled: Boolean(logicalName && entitySetName),
    staleTime: 10 * 60 * 1000,
    queryFn: async () => {
      if (!isPowerAppsRuntime()) return fallbackColumns[logicalName] ?? [];
      try {
        const metadata = await loadEntityMetadata(logicalName);
        const columns = mapDataverseColumns(metadata);
        return columns.length ? columns : fallbackColumns[logicalName] ?? [];
      } catch {
        return fallbackColumns[logicalName] ?? [];
      }
    },
  });
}

export function useDataverseRelationships(logicalName: string, entitySetName: string) {
  return useQuery({
    queryKey: ["dataverse-relationships", logicalName, entitySetName],
    enabled: Boolean(logicalName && entitySetName),
    staleTime: 10 * 60 * 1000,
    queryFn: async () => {
      if (!isPowerAppsRuntime()) return fallbackRelationships[logicalName] ?? [];
      try {
        const metadata = await loadEntityMetadata(logicalName);
        const relationships = mapDataverseRelationships(metadata, logicalName);
        return relationships.length ? relationships : fallbackRelationships[logicalName] ?? [];
      } catch {
        return fallbackRelationships[logicalName] ?? [];
      }
    },
  });
}

export interface DataversePreviewInput {
  logicalName: string;
  entitySetName: string;
  fetchXml: string;
  limit?: number;
}

/**
 * The Dataverse connector paginates ListRecords by injecting `page`/`count`
 * into the FetchXML, and Dataverse rejects `top` combined with paging
 * ("The top attribute can't be specified with paging attribute page").
 * For previews we therefore strip every limit/paging attribute from the
 * `<fetch>` element and apply the limit with the OData `$top` parameter,
 * which Dataverse accepts alongside FetchXML.
 */
export function prepareFetchXmlForPreview(fetchXml: string) {
  return fetchXml.replace(/<fetch\b([^>]*)>/i, (_match, attributes: string) => {
    const cleaned = attributes
      .replace(/\s+(top|page|count|paging-cookie|returntotalrecordcount)=(['"])[^'"]*\2/gi, "")
      .replace(/\s+$/, "");
    return `<fetch${cleaned}>`;
  });
}

function unwrapResult(result: unknown, fallbackMessage: string) {
  const outcome = result as
    | { success?: boolean; error?: { message?: string; status?: number } }
    | null;
  if (outcome && outcome.success === false) {
    const message = outcome.error?.message?.trim() || fallbackMessage;
    throw new Error(
      outcome.error?.status ? `${message} (HTTP ${outcome.error.status})` : message,
    );
  }
}

export function useDataverseRecordPreview() {
  return useMutation({
    mutationFn: async ({ logicalName, entitySetName, fetchXml, limit = 10 }: DataversePreviewInput) => {
      if (!fetchXml.trim().startsWith("<fetch")) throw new Error("Enter valid FetchXML before previewing records.");
      if (!isPowerAppsRuntime()) {
        const samples: Record<string, Array<Record<string, unknown>>> = {
          account: [
            { accountid: "mock-account-1", name: "Contoso Retail", revenue: 8400000 },
            { accountid: "mock-account-2", name: "Northwind Health", revenue: 5100000 },
            { accountid: "mock-account-3", name: "Fabrikam Energy", revenue: 12700000 },
          ],
          contact: [
            { contactid: "mock-contact-1", fullname: "Adele Vance", emailaddress1: "adele@example.com" },
            { contactid: "mock-contact-2", fullname: "Nestor Wilke", emailaddress1: "nestor@example.com" },
          ],
        };
        return (samples[logicalName] ?? []).slice(0, limit);
      }
      return runFetchXml(entitySetName, fetchXml, limit);
    },
  });
}
