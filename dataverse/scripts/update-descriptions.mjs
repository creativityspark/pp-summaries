import { PublicClientApplication } from "@azure/msal-node";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const API_VERSION = "v9.2";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const label = (text) => ({
  "@odata.type": "Microsoft.Dynamics.CRM.Label",
  LocalizedLabels: [
    {
      "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
      Label: text,
      LanguageCode: 1033
    }
  ]
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function getToken() {
  const pca = new PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`
    }
  });
  const result = await pca.acquireTokenByDeviceCode({
    scopes: [`${ENV_URL}/.default`],
    deviceCodeCallback: (r) => {
      console.log("\n=========================================");
      console.log("DEVICE CODE AUTHENTICATION");
      console.log("=========================================");
      console.log(r.message);
      console.log("=========================================\n");
    }
  });
  return result.accessToken;
}

async function api(token, method, urlPath, body, extraHeaders = {}) {
  const url = `${ENV_URL}/api/data/${API_VERSION}${urlPath}`;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "MSCRM.MergeLabels": "true",
      ...extraHeaders
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${method} ${urlPath} -> ${res.status} ${res.statusText}\n${text}`);
  }
  if (res.status === 204) return null;
  const ct = res.headers.get("content-type") || "";
  return ct.includes("application/json") ? res.json() : res.text();
}

async function apiWithRetry(token, method, urlPath, body, extraHeaders = {}) {
  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await api(token, method, urlPath, body, extraHeaders);
    } catch (err) {
      const transient = /\b(400|429|500|502|503|504)\b/.test(err.message);
      if (attempt < maxAttempts && transient) {
        const wait = 1500 * attempt;
        console.log(`    [retry ${attempt}/${maxAttempts - 1}] after ${wait}ms`);
        await sleep(wait);
        continue;
      }
      throw err;
    }
  }
}

function getCurrentLabel(localized) {
  if (!localized) return "";
  const obj = localized.UserLocalizedLabel || localized.LocalizedLabels?.[0];
  return obj?.Label || "";
}

async function main() {
  const cfg = JSON.parse(await readFile(path.join(__dirname, "tables.json"), "utf-8"));

  console.log("Authenticating...");
  const token = await getToken();
  console.log("Authenticated.\n");

  let updated = 0;
  let skipped = 0;

  for (const table of cfg.tables) {
    const logicalName = table.schemaName.toLowerCase();
    console.log(`\n=== ${table.schemaName} ===`);

    const current = await api(
      token,
      "GET",
      `/EntityDefinitions(LogicalName='${logicalName}')?$select=MetadataId,Description`
    );
    const currentDesc = getCurrentLabel(current.Description);

    if (currentDesc !== table.description) {
      console.log(`  [update] table description`);
      await apiWithRetry(
        token,
        "PUT",
        `/EntityDefinitions(${current.MetadataId})`,
        {
          "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
          MetadataId: current.MetadataId,
          LogicalName: logicalName,
          SchemaName: table.schemaName,
          Description: label(table.description)
        },
        { "If-Match": "*" }
      );
      updated++;
      await sleep(400);
    } else {
      console.log(`  [skip] table description unchanged`);
      skipped++;
    }

    const primary = table.primaryAttribute;
    const primaryLogical = primary.schemaName.toLowerCase();
    const allAttrs = [primary, ...table.attributes];

    for (const attr of allAttrs) {
      const attrLogical = attr.schemaName.toLowerCase();

      if (attr.type === "Lookup") {
        console.log(`  [skip] ${attr.schemaName} - lookup updates not yet supported`);
        skipped++;
        continue;
      }

      let currentAttr;
      try {
        currentAttr = await api(
          token,
          "GET",
          `/EntityDefinitions(LogicalName='${logicalName}')/Attributes(LogicalName='${attrLogical}')?$select=MetadataId,Description`
        );
      } catch (err) {
        console.log(`  [warn] ${attr.schemaName} not found - ${err.message.split("\n")[0]}`);
        continue;
      }

      const currentAttrDesc = getCurrentLabel(currentAttr.Description);
      if (currentAttrDesc !== attr.description) {
        console.log(`  [update] ${attr.schemaName}`);
        const attrType = attr.type === undefined ? "String" : attr.type;
        const odataType = {
          String: "Microsoft.Dynamics.CRM.StringAttributeMetadata",
          Memo: "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
          Integer: "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
          Decimal: "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
          DateTime: "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
          Boolean: "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
          Picklist: "Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
        }[attrType];
        await apiWithRetry(
          token,
          "PUT",
          `/EntityDefinitions(LogicalName='${logicalName}')/Attributes(${currentAttr.MetadataId})`,
          {
            "@odata.type": odataType,
            MetadataId: currentAttr.MetadataId,
            LogicalName: attrLogical,
            SchemaName: attr.schemaName,
            Description: label(attr.description)
          }
        );
        updated++;
        await sleep(300);
      } else {
        skipped++;
      }
    }
  }

  console.log(`\n=== Summary ===`);
  console.log(`Updated: ${updated}`);
  console.log(`Skipped (unchanged or unsupported): ${skipped}`);
  console.log("\nDone.");
}

main().catch((err) => {
  console.error("\nFATAL:", err.message);
  process.exit(1);
});
