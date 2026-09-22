import { useQuery } from "@tanstack/react-query";
import { getClient } from "@microsoft/power-apps/data";
import { dataSourcesInfo } from "../../.power/schemas/appschemas/dataSourcesInfo";
import { bucketFor, type AttributeBucket } from "@/lib/fetchxml";

export interface EntityAttribute {
  logicalName: string;
  displayName: string;
  /** Raw AttributeType reported by Dataverse, e.g. "String", "Picklist". */
  rawType: string;
  /** High-level bucket used by the UI to pick operators and value pickers. */
  bucket: AttributeBucket;
  /** True for system/internal columns the user rarely wants in a filter. */
  isSystem: boolean;
}

/**
 * Map of logicalName → entitySetName for entities the SDK can call.
 * Same list as useFetchXmlPreview — keep in sync.
 */
const ENTITY_SET_NAME: Record<string, string> = {
  account: "accounts",
  contact: "contacts",
  csp_contactlogentry: "csp_contactlogentries",
  csp_nextbestaction: "csp_nextbestactions",
  csp_accounthealth: "csp_accounthealths",
  csp_aisummarycache: "csp_aisummarycaches",
  csp_weeklyrecap: "csp_weeklyrecaps",
  csp_aiusage: "csp_aiusages",
  csp_aiprompt: "csp_aiprompts",
  csp_aisummaryconfig: "csp_aisummaryconfigs",
  savedquery: "savedqueries",
  userquery: "userqueries",
};

/**
 * Logical names that are noisy/system and should be hidden from the
 * primary attribute picker. The user can still type them manually.
 */
const HIDDEN_ATTRIBUTES = new Set([
  "createdbyname",
  "createdbyyominame",
  "createdonbehalfby",
  "createdonbehalfbyname",
  "createdonbehalfbyyominame",
  "modifiedbyname",
  "modifiedbyyominame",
  "modifiedonbehalfby",
  "modifiedonbehalfbyname",
  "modifiedonbehalfbyyominame",
  "owningbusinessunit",
  "owningbusinessunitname",
  "owningteam",
  "owninguser",
  "owneridname",
  "owneridyominame",
  "owneridtype",
  "owneridsname",
  "owneridtypename",
  "importsequencenumber",
  "overriddencreatedon",
  "timezoneruleversionnumber",
  "utcconversiontimezonecode",
  "versionnumber",
  "statecodename",
  "statuscodename",
]);

const SYSTEM_PREFIXES = ["_", "yomi"];

function isSystemAttribute(logicalName: string): boolean {
  if (HIDDEN_ATTRIBUTES.has(logicalName)) return true;
  if (SYSTEM_PREFIXES.some((p) => logicalName.startsWith(p))) return true;
  return false;
}

const sharedClient = getClient(dataSourcesInfo);

/* ─── Well-known attribute lists ─────────────────────────────────── */

/**
 * Curated attribute lists for the entities this demo knows ahead of time.
 * Used as a fallback when the Dataverse metadata SDK call returns nothing.
 * Includes only the columns that are useful in a filter — not the full ~80.
 *
 * For account/contact this is the standard Dataverse out-of-the-box schema.
 * For csp_* it's the schema defined in dataverse/scripts/tables.json plus
 * the system fields the platform adds automatically (statecode, modifiedon,
 * ownerid, etc.).
 */
function attr(
  logicalName: string,
  displayName: string,
  rawType: string,
  isSystem = false
): EntityAttribute {
  return {
    logicalName,
    displayName,
    rawType,
    bucket: bucketFor(rawType),
    isSystem,
  };
}

const SYSTEM_FIELDS: EntityAttribute[] = [
  attr("createdon", "Created On", "DateTime", true),
  attr("modifiedon", "Modified On", "DateTime", true),
  attr("createdby", "Created By", "Lookup", true),
  attr("modifiedby", "Modified By", "Lookup", true),
  attr("ownerid", "Owner", "Owner", true),
  attr("statecode", "Status", "State", true),
  attr("statuscode", "Status Reason", "Status", true),
];

const WELL_KNOWN_ATTRIBUTES: Record<string, EntityAttribute[]> = {
  account: [
    attr("name", "Account Name", "String"),
    attr("accountnumber", "Account Number", "String"),
    attr("address1_city", "City", "String"),
    attr("address1_country", "Country", "String"),
    attr("address1_postalcode", "Postal Code", "String"),
    attr("address1_stateorprovince", "State / Province", "String"),
    attr("address1_line1", "Street", "String"),
    attr("telephone1", "Main Phone", "String"),
    attr("emailaddress1", "Email", "String"),
    attr("websiteurl", "Website", "String"),
    attr("description", "Description", "Memo"),
    attr("revenue", "Annual Revenue", "Money"),
    attr("numberofemployees", "Number of Employees", "Integer"),
    attr("industrycode", "Industry", "Picklist"),
    attr("customertypecode", "Relationship Type", "Picklist"),
    attr("accountratingcode", "Account Rating", "Picklist"),
    attr("creditlimit", "Credit Limit", "Money"),
    attr("creditonhold", "Credit Hold", "Boolean"),
    attr("preferredcontactmethodcode", "Preferred Contact Method", "Picklist"),
    attr("parentaccountid", "Parent Account", "Lookup"),
    attr("primarycontactid", "Primary Contact", "Lookup"),
    attr("csp_isdemo", "Is Demo Record", "Boolean"),
    ...SYSTEM_FIELDS,
  ],
  contact: [
    attr("fullname", "Full Name", "String"),
    attr("firstname", "First Name", "String"),
    attr("lastname", "Last Name", "String"),
    attr("jobtitle", "Job Title", "String"),
    attr("emailaddress1", "Email", "String"),
    attr("telephone1", "Business Phone", "String"),
    attr("mobilephone", "Mobile Phone", "String"),
    attr("address1_city", "City", "String"),
    attr("address1_country", "Country", "String"),
    attr("description", "Description", "Memo"),
    attr("parentcustomerid", "Company", "Customer"),
    attr("preferredcontactmethodcode", "Preferred Contact Method", "Picklist"),
    attr("birthdate", "Birthday", "DateTime"),
    attr("csp_isdemo", "Is Demo Record", "Boolean"),
    ...SYSTEM_FIELDS,
  ],
  csp_contactlogentry: [
    attr("csp_name", "Name", "String"),
    attr("csp_rawinput", "Raw Input", "Memo"),
    attr("csp_summary", "Summary", "Memo"),
    attr("csp_sentiment", "Sentiment", "Picklist"),
    attr("csp_interactiontype", "Interaction Type", "Picklist"),
    attr("csp_competitorsmentioned", "Competitors Mentioned", "String"),
    attr("csp_productsmentioned", "Products Mentioned", "String"),
    attr("csp_proposednextaction", "Proposed Next Action", "String"),
    attr("csp_interactiondate", "Interaction Date", "DateTime"),
    attr("csp_account", "Account", "Lookup"),
    attr("csp_contact", "Contact", "Lookup"),
    ...SYSTEM_FIELDS,
  ],
  csp_nextbestaction: [
    attr("csp_name", "Name", "String"),
    attr("csp_description", "Description", "Memo"),
    attr("csp_justification", "Justification", "Memo"),
    attr("csp_actiontype", "Action Type", "Picklist"),
    attr("csp_priority", "Priority", "Picklist"),
    attr("csp_status", "Status", "Picklist"),
    attr("csp_generatedon", "Generated On", "DateTime"),
    attr("csp_executedon", "Executed On", "DateTime"),
    attr("csp_account", "Account", "Lookup"),
    ...SYSTEM_FIELDS,
  ],
  csp_accounthealth: [
    attr("csp_name", "Name", "String"),
    attr("csp_score", "Score", "Integer"),
    attr("csp_band", "Band", "Picklist"),
    attr("csp_signals", "Signals", "Memo"),
    attr("csp_calculatedon", "Calculated On", "DateTime"),
    attr("csp_account", "Account", "Lookup"),
    ...SYSTEM_FIELDS,
  ],
  csp_aisummarycache: [
    attr("csp_name", "Name", "String"),
    attr("csp_content", "Content", "Memo"),
    attr("csp_summarytype", "Summary Type", "Picklist"),
    attr("csp_generatedon", "Generated On", "DateTime"),
    attr("csp_tokensused", "Tokens Used", "Integer"),
    attr("csp_promptversion", "Prompt Version", "String"),
    attr("csp_targetentity", "Target Entity", "String"),
    attr("csp_targetrecordid", "Target Record Id", "String"),
    attr("csp_account", "Account", "Lookup"),
    attr("csp_summaryconfig", "Summary Config", "Lookup"),
    attr("csp_prompt", "Prompt", "Lookup"),
    ...SYSTEM_FIELDS,
  ],
  csp_weeklyrecap: [
    attr("csp_name", "Name", "String"),
    attr("csp_weeknumber", "Week Number", "Integer"),
    attr("csp_year", "Year", "Integer"),
    attr("csp_blogtitle", "Blog Title", "String"),
    attr("csp_blogcontent", "Blog Content", "Memo"),
    attr("csp_audiourl", "Audio URL", "String"),
    attr("csp_audioduration", "Audio Duration", "Integer"),
    attr("csp_coverurl", "Cover URL", "String"),
    attr("csp_highlights", "Highlights", "Memo"),
    attr("csp_insight", "Insight", "Memo"),
    attr("csp_tokensused", "Tokens Used", "Integer"),
    ...SYSTEM_FIELDS,
  ],
  csp_aiusage: [
    attr("csp_name", "Name", "String"),
    attr("csp_promptname", "Prompt Name", "String"),
    attr("csp_tokensinput", "Tokens Input", "Integer"),
    attr("csp_tokensoutput", "Tokens Output", "Integer"),
    attr("csp_estimatedcost", "Estimated Cost", "Decimal"),
    attr("csp_latencyms", "Latency (ms)", "Integer"),
    attr("csp_model", "Model", "String"),
    attr("csp_timestamp", "Timestamp", "DateTime"),
    attr("csp_targetentity", "Target Entity", "String"),
    attr("csp_targetrecordid", "Target Record Id", "String"),
    attr("csp_account", "Account", "Lookup"),
    attr("csp_user", "User", "Lookup"),
    attr("csp_promptlookup", "Prompt", "Lookup"),
    attr("csp_summaryconfig", "Summary Config", "Lookup"),
    ...SYSTEM_FIELDS,
  ],
  csp_aiprompt: [
    attr("csp_name", "Name", "String"),
    attr("csp_promptkey", "Prompt Key", "String"),
    attr("csp_description", "Description", "String"),
    attr("csp_systemmessage", "System Message", "Memo"),
    attr("csp_content", "Prompt Content", "Memo"),
    attr("csp_expectedvariables", "Expected Variables", "Memo"),
    attr("csp_outputformat", "Output Format", "Picklist"),
    attr("csp_outputschema", "Output Schema", "Memo"),
    attr("csp_model", "Model", "String"),
    attr("csp_deploymentname", "Deployment Name", "String"),
    attr("csp_version", "Version", "String"),
    attr("csp_enabled", "Enabled", "Boolean"),
    attr("csp_temperature", "Temperature", "Decimal"),
    attr("csp_maxtokens", "Max Tokens", "Integer"),
    attr("csp_lastrunon", "Last Run On", "DateTime"),
    attr("csp_tokenslastrun", "Tokens Last Run", "Integer"),
    ...SYSTEM_FIELDS,
  ],
  csp_aisummaryconfig: [
    attr("csp_name", "Name", "String"),
    attr("csp_description", "Description", "Memo"),
    attr("csp_cadence", "Cadence", "Picklist"),
    attr("csp_cronexpression", "Cron Expression", "String"),
    attr("csp_mode", "Mode", "Picklist"),
    attr("csp_prompt", "Prompt", "Lookup"),
    attr("csp_targetentity", "Target Entity", "String"),
    attr("csp_querymode", "Query Mode", "Picklist"),
    attr("csp_systemviewid", "System View Id", "String"),
    attr("csp_userviewid", "User View Id", "String"),
    attr("csp_fetchxml", "FetchXML", "Memo"),
    attr("csp_relatedfetchxml", "Related FetchXML", "Memo"),
    attr("csp_maxrecords", "Max Records", "Integer"),
    attr("csp_language", "Language", "String"),
    attr("csp_outputdestination", "Output Destination", "Picklist"),
    attr("csp_outputcontext", "Output Context", "Memo"),
    attr("csp_notifyrecipients", "Notify Recipients", "Memo"),
    attr("csp_enabled", "Enabled", "Boolean"),
    attr("csp_nextrun", "Next Run", "DateTime"),
    attr("csp_lastrun", "Last Run", "DateTime"),
    attr("csp_lastrunstatus", "Last Run Status", "Picklist"),
    attr("csp_lastrunerror", "Last Run Error", "Memo"),
    attr("csp_lastrundurationms", "Last Run Duration (ms)", "Integer"),
    ...SYSTEM_FIELDS,
  ],
  savedquery: [
    attr("name", "Name", "String"),
    attr("description", "Description", "Memo"),
    attr("returnedtypecode", "Entity", "String"),
    attr("querytype", "Query Type", "Integer"),
    attr("isdefault", "Is Default", "Boolean"),
    attr("fetchxml", "FetchXML", "Memo"),
    ...SYSTEM_FIELDS,
  ],
  userquery: [
    attr("name", "Name", "String"),
    attr("description", "Description", "Memo"),
    attr("returnedtypecode", "Entity", "String"),
    attr("querytype", "Query Type", "Integer"),
    attr("fetchxml", "FetchXML", "Memo"),
    ...SYSTEM_FIELDS,
  ],
};

/* ─── SDK metadata extractor ─────────────────────────────────────── */

/**
 * Try to extract attributes from whatever shape the SDK happens to return
 * from getEntityMetadata. The shape varies across SDK versions and tenants,
 * so we look in a few well-known places before giving up.
 */
function extractAttributes(meta: unknown): Array<Record<string, unknown>> {
  if (!meta || typeof meta !== "object") return [];
  const candidates = [
    (meta as Record<string, unknown>).attributes,
    (meta as Record<string, unknown>).Attributes,
    ((meta as Record<string, unknown>).data as Record<string, unknown> | undefined)
      ?.attributes,
    ((meta as Record<string, unknown>).data as Record<string, unknown> | undefined)
      ?.Attributes,
  ];
  for (const c of candidates) {
    if (Array.isArray(c)) return c as Array<Record<string, unknown>>;
  }
  return [];
}

function readLabel(raw: Record<string, unknown>): string {
  const dn = raw.DisplayName ?? raw.displayName;
  if (typeof dn === "string") return dn;
  if (dn && typeof dn === "object") {
    const obj = dn as Record<string, unknown>;
    const userLocal = obj.UserLocalizedLabel as Record<string, unknown> | undefined;
    if (userLocal && typeof userLocal.Label === "string") return userLocal.Label;
    if (typeof obj.Label === "string") return obj.Label;
    const localized = obj.LocalizedLabels;
    if (Array.isArray(localized) && localized.length > 0) {
      const first = localized[0] as Record<string, unknown>;
      if (typeof first.Label === "string") return first.Label;
    }
  }
  return "";
}

function mapMetaToAttribute(raw: Record<string, unknown>): EntityAttribute | null {
  const logicalName =
    (raw.logicalName as string) ?? (raw.LogicalName as string) ?? "";
  if (!logicalName) return null;
  const rawType =
    (raw.attributeType as string) ??
    (raw.AttributeType as string) ??
    (raw.type as string) ??
    "Unknown";
  const displayName = readLabel(raw) || logicalName;
  return {
    logicalName,
    displayName,
    rawType,
    bucket: bucketFor(rawType),
    isSystem: isSystemAttribute(logicalName),
  };
}

/**
 * Merge two attribute lists, preferring the SDK one when a logical name
 * appears in both. Adds any well-known attribute that the SDK didn't return.
 */
function mergeAttributes(
  fromSdk: EntityAttribute[],
  fromWellKnown: EntityAttribute[]
): EntityAttribute[] {
  const map = new Map<string, EntityAttribute>();
  for (const a of fromWellKnown) map.set(a.logicalName, a);
  for (const a of fromSdk) map.set(a.logicalName, a);
  return Array.from(map.values());
}

/**
 * Fetch the attribute list for a given target entity.
 *   1. Tries the Dataverse metadata SDK call.
 *   2. If that returns nothing, falls back to a curated WELL_KNOWN list.
 *   3. If still nothing, returns an empty array (UI falls back to typing).
 *
 * The SDK response shape is logged to the console (once per entity) to
 * help debug the parser when running against new tenants/SDK versions.
 */
export function useEntityAttributes(entityName: string | undefined) {
  return useQuery({
    queryKey: ["entity-attributes", entityName],
    enabled: !!entityName,
    staleTime: 5 * 60 * 1000,
    queryFn: async (): Promise<EntityAttribute[]> => {
      if (!entityName) return [];
      const wellKnown = WELL_KNOWN_ATTRIBUTES[entityName] ?? [];
      const entitySet = ENTITY_SET_NAME[entityName];

      let fromSdk: EntityAttribute[] = [];
      if (entitySet) {
        try {
          const res = await sharedClient.executeAsync({
            dataverseRequest: {
              action: "getEntityMetadata",
              parameters: {
                tableName: entitySet,
                options: {},
              },
            },
          });
          // Log raw response so we can iterate on the parser when needed.
          console.info(
            `[useEntityAttributes] raw metadata for "${entityName}":`,
            res
          );
          const raw = extractAttributes(res?.data ?? res);
          fromSdk = raw
            .map(mapMetaToAttribute)
            .filter((a): a is EntityAttribute => !!a);
          if (fromSdk.length === 0) {
            console.warn(
              `[useEntityAttributes] SDK returned 0 attributes for "${entityName}", using well-known fallback (${wellKnown.length} attrs)`
            );
          }
        } catch (err) {
          console.warn(
            `[useEntityAttributes] metadata fetch failed for "${entityName}", using well-known fallback:`,
            err
          );
        }
      }

      const merged = mergeAttributes(fromSdk, wellKnown);
      merged.sort((a, b) => {
        if (a.isSystem !== b.isSystem) return a.isSystem ? 1 : -1;
        return a.displayName.localeCompare(b.displayName);
      });
      return merged;
    },
  });
}
