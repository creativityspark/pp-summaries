/**
 * Remove all synthetic data from the demo environment.
 *
 * - Truncates all custom csp_* tables (they only ever hold demo data).
 * - Deletes records from standard tables (account, contact) where csp_isdemo = true.
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

const CUSTOM_TABLES = [
  { entitySet: "csp_aiusages", logical: "csp_aiusage", primaryKey: "csp_aiusageid" },
  { entitySet: "csp_userfeedbacks", logical: "csp_userfeedback", primaryKey: "csp_userfeedbackid" },
  { entitySet: "csp_aisummarycaches", logical: "csp_aisummarycache", primaryKey: "csp_aisummarycacheid" },
  { entitySet: "csp_accounthealths", logical: "csp_accounthealth", primaryKey: "csp_accounthealthid" },
  { entitySet: "csp_nextbestactions", logical: "csp_nextbestaction", primaryKey: "csp_nextbestactionid" },
  { entitySet: "csp_contactlogentries", logical: "csp_contactlogentry", primaryKey: "csp_contactlogentryid" },
  { entitySet: "csp_weeklyrecaps", logical: "csp_weeklyrecap", primaryKey: "csp_weeklyrecapid" }
];

const STANDARD_TABLES = [
  { entitySet: "contacts", primaryKey: "contactid", filter: "csp_isdemo eq true" },
  { entitySet: "accounts", primaryKey: "accountid", filter: "csp_isdemo eq true" }
];

const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`
  }
});

const token = (await pca.acquireTokenByDeviceCode({
  scopes: [`${ENV_URL}/.default`],
  deviceCodeCallback: (r) => {
    console.log("\n=========================================");
    console.log("DEVICE CODE AUTHENTICATION");
    console.log("=========================================");
    console.log(r.message);
    console.log("=========================================\n");
  }
})).accessToken;
console.log("Authenticated.\n");

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function listIds(entitySet, primaryKey, filter) {
  const all = [];
  let url = `${ENV_URL}/api/data/v9.2/${entitySet}?$select=${primaryKey}${filter ? `&$filter=${encodeURIComponent(filter)}` : ""}`;
  while (url) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        Prefer: 'odata.maxpagesize=500'
      }
    });
    if (!res.ok) throw new Error(`List ${entitySet}: ${res.status}\n${await res.text()}`);
    const data = await res.json();
    for (const r of data.value) all.push(r[primaryKey]);
    url = data["@odata.nextLink"] ?? null;
  }
  return all;
}

async function deleteRecord(entitySet, id) {
  const res = await fetch(`${ENV_URL}/api/data/v9.2/${entitySet}(${id})`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok && res.status !== 404) {
    throw new Error(`DELETE ${entitySet}(${id}): ${res.status}\n${await res.text()}`);
  }
}

// truncate custom tables first (they only ever hold demo)
for (const t of CUSTOM_TABLES) {
  const ids = await listIds(t.entitySet, t.primaryKey);
  console.log(`[${t.entitySet}] ${ids.length} records to delete`);
  let i = 0;
  for (const id of ids) {
    try {
      await deleteRecord(t.entitySet, id);
    } catch (err) {
      console.log(`  [fail] ${id}: ${err.message.split("\n")[0]}`);
    }
    i++;
    if (i % 25 === 0) {
      console.log(`  ${i}/${ids.length}`);
      await sleep(100);
    }
  }
  console.log(`  done: ${i} deleted\n`);
}

// then delete demo records from standard tables
for (const t of STANDARD_TABLES) {
  const ids = await listIds(t.entitySet, t.primaryKey, t.filter);
  console.log(`[${t.entitySet}] ${ids.length} demo records to delete`);
  let i = 0;
  for (const id of ids) {
    try {
      await deleteRecord(t.entitySet, id);
    } catch (err) {
      console.log(`  [fail] ${id}: ${err.message.split("\n")[0]}`);
    }
    i++;
    if (i % 25 === 0) {
      console.log(`  ${i}/${ids.length}`);
      await sleep(100);
    }
  }
  console.log(`  done: ${i} deleted\n`);
}

console.log("Done.");
