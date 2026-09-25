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
const connectorName = "commondataserviceforapps";

const fallbackTables: DataverseTableMetadata[] = [
  { logicalName: "account", entitySetName: "accounts", primaryIdAttribute: "accountid", displayName: "Accounts" },
  { logicalName: "contact", entitySetName: "contacts", primaryIdAttribute: "contactid", displayName: "Contacts" },
  { logicalName: "incident", entitySetName: "incidents", primaryIdAttribute: "incidentid", displayName: "Cases" },
  { logicalName: "opportunity", entitySetName: "opportunities", primaryIdAttribute: "opportunityid", displayName: "Opportunities" },
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

function isPowerAppsRuntime() {
  if (typeof window === "undefined") return false;
  return window.parent !== window || window.location.hostname.endsWith("powerapps.com");
}

async function loadEntityMetadata(logicalName: string) {
  const escapedName = logicalName.replaceAll("'", "''");
  const result = await client.executeAsync<Record<string, unknown>, unknown>({
    connectorOperation: {
      tableName: connectorName,
      operationName: "ListRecords",
      parameters: {
        entityName: "EntityDefinitions",
        $select: "LogicalName,EntitySetName,PrimaryIdAttribute,DisplayCollectionName",
        $filter: `LogicalName eq '${escapedName}'`,
        $expand: "Attributes($select=LogicalName,DisplayName,AttributeType,IsValidForRead,IsValidForUpdate),OneToManyRelationships($select=SchemaName,ReferencedEntity,ReferencedAttribute,ReferencingEntity,ReferencingAttribute),ManyToOneRelationships($select=SchemaName,ReferencedEntity,ReferencedAttribute,ReferencingEntity,ReferencingAttribute)",
        $top: 1,
      },
    },
  });
  return mapDataverseRecords(result)[0] ?? {};
}

export function useDataverseTables() {
  return useQuery({
    queryKey: ["dataverse-table-catalog"],
    staleTime: 10 * 60 * 1000,
    queryFn: async () => {
      if (!isPowerAppsRuntime()) return fallbackTables;
      try {
        const result = await client.executeAsync<Record<string, unknown>, unknown>({
          connectorOperation: {
            tableName: connectorName,
            operationName: "ListRecords",
            parameters: { entityName: "EntityDefinitions", $select: "LogicalName,EntitySetName,PrimaryIdAttribute,DisplayCollectionName", $top: 5000 },
          },
        });
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

export function limitFetchXmlForPreview(fetchXml: string, limit: number) {
  const safeLimit = Math.max(1, Math.floor(limit));
  return fetchXml.replace(/<fetch\b([^>]*)>/i, (_match, attributes: string) => {
    const withoutTop = attributes.replace(/\s+top=(['"])[^'"]*\1/i, "");
    return `<fetch${withoutTop} top="${safeLimit}">`;
  });
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
      const result = await client.executeAsync<Record<string, unknown>, unknown>({
        connectorOperation: {
          tableName: connectorName,
          operationName: "ListRecords",
          parameters: { entityName: entitySetName, fetchXml: limitFetchXmlForPreview(fetchXml, limit) },
        },
      });
      return mapDataverseRecords(result).slice(0, limit);
    },
  });
}
