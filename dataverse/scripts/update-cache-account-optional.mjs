/**
 * One-off migration: change csp_aisummarycache.csp_account RequiredLevel
 * from ApplicationRequired to None.
 *
 * Needed because create-tables.mjs only adds new attributes; it does not
 * modify the metadata of attributes that already exist. After the model
 * pivot to generic target records (csp_TargetEntity + csp_TargetRecordId),
 * csp_Account must be optional so the cache can hold summaries for any
 * Dataverse table, not just account.
 *
 * Safe to run multiple times — exits early if the attribute is already None.
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const API_VERSION = "v9.2";

const TARGET_ENTITY = "csp_aisummarycache";
const TARGET_ATTRIBUTE = "csp_account";

const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
});

const token = (
  await pca.acquireTokenByDeviceCode({
    scopes: [`${ENV_URL}/.default`],
    deviceCodeCallback: (r) => {
      console.log("\n=========================================");
      console.log("DEVICE CODE AUTHENTICATION");
      console.log("=========================================");
      console.log(r.message);
      console.log("=========================================\n");
    },
  })
).accessToken;
console.log("Authenticated.\n");

async function api(method, urlPath, body, extraHeaders = {}) {
  const url = `${ENV_URL}/api/data/${API_VERSION}${urlPath}`;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "MSCRM.SolutionUniqueName": "BizzSummit2026",
      ...extraHeaders,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${method} ${urlPath} -> ${res.status} ${res.statusText}\n${text}`);
  }
  if (res.status === 204) return null;
  const ct = res.headers.get("content-type") || "";
  return ct.includes("application/json") ? res.json() : res.text();
}

console.log(`Fetching current metadata for ${TARGET_ENTITY}.${TARGET_ATTRIBUTE}...`);
const current = await api(
  "GET",
  `/EntityDefinitions(LogicalName='${TARGET_ENTITY}')/Attributes(LogicalName='${TARGET_ATTRIBUTE}')`
);

console.log(`Current RequiredLevel: ${current.RequiredLevel?.Value ?? "(none)"}`);

if (current.RequiredLevel?.Value === "None") {
  console.log("Already None. Nothing to update.\n");
  process.exit(0);
}

const body = {
  "@odata.type": "Microsoft.Dynamics.CRM.LookupAttributeMetadata",
  AttributeType: "Lookup",
  AttributeTypeName: { Value: "LookupType" },
  SchemaName: current.SchemaName,
  LogicalName: current.LogicalName,
  MetadataId: current.MetadataId,
  RequiredLevel: { Value: "None", CanBeChanged: true },
};

console.log("Updating RequiredLevel to None...");
await api(
  "PUT",
  `/EntityDefinitions(LogicalName='${TARGET_ENTITY}')/Attributes(LogicalName='${TARGET_ATTRIBUTE}')`,
  body,
  { "MSCRM.MergeLabels": "true" }
);

console.log(`✓ ${TARGET_ENTITY}.${TARGET_ATTRIBUTE} RequiredLevel set to None.\n`);
