import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const API_VERSION = "v9.2";

const TABLES_TO_DELETE = process.argv.slice(2);

if (TABLES_TO_DELETE.length === 0) {
  console.error("Usage: node delete-tables.mjs <table_logical_name> [<more>...]");
  process.exit(1);
}

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

async function api(token, method, urlPath) {
  const res = await fetch(`${ENV_URL}/api/data/${API_VERSION}${urlPath}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0"
    }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${method} ${urlPath} -> ${res.status} ${res.statusText}\n${text}`);
  }
  if (res.status === 204) return null;
  const ct = res.headers.get("content-type") || "";
  return ct.includes("application/json") ? res.json() : res.text();
}

const token = await getToken();
console.log("Authenticated.\n");

for (const logicalName of TABLES_TO_DELETE) {
  try {
    const ent = await api(token, "GET", `/EntityDefinitions(LogicalName='${logicalName}')?$select=MetadataId,LogicalName`);
    console.log(`Deleting ${logicalName} (${ent.MetadataId})...`);
    await api(token, "DELETE", `/EntityDefinitions(${ent.MetadataId})`);
    console.log(`  [ok] ${logicalName} deleted`);
  } catch (err) {
    console.log(`  [fail] ${logicalName}: ${err.message.split("\n")[0]}`);
  }
}

console.log("\nDone.");
