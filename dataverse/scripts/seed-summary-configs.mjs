/**
 * Seed 3 example AI Summary Configs into csp_aisummaryconfigs.
 *
 * Each config is a declarative recipe for the run-summary-config flow:
 *   - which records to summarise (FetchXML query)
 *   - which prompt to use (lookup to csp_aiprompt by promptkey)
 *   - cadence + mode (PerRecord or Aggregate)
 *   - output destination
 *
 * The seed demonstrates three distinct combinations to power the
 * "three configs, same engine" demo moment on stage.
 *
 * Re-runs are idempotent: deletes all existing demo configs first,
 * then creates the three from scratch. Marked as demo via the config's
 * description (no csp_isdemo column on csp_aisummaryconfig).
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

// Picklist values — keep in sync with tables.json
const CADENCE = {
  OnDemand: 100000000,
  Daily: 100000001,
  Weekly: 100000002,
  Monthly: 100000003,
  Custom: 100000004,
};
const MODE = {
  PerRecord: 100000000,
  Aggregate: 100000001,
};
const QUERY_MODE = {
  SystemView: 100000000,
  UserView: 100000001,
  CustomFetchXML: 100000002,
};
const OUTPUT_DESTINATION = {
  Cache: 100000000,
  CacheAndEmail: 100000001,
  CacheAndTeams: 100000002,
  CacheEmailAndTeams: 100000003,
};

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

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function listAll(entitySet, select) {
  const all = [];
  let url = `${ENV_URL}/api/data/v9.2/${entitySet}?$select=${select}`;
  while (url) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        Prefer: "odata.maxpagesize=500",
      },
    });
    if (!res.ok) throw new Error(`List ${entitySet}: ${res.status}\n${await res.text()}`);
    const data = await res.json();
    for (const r of data.value) all.push(r);
    url = data["@odata.nextLink"] ?? null;
  }
  return all;
}

async function deleteRecord(entitySet, id) {
  const res = await fetch(`${ENV_URL}/api/data/v9.2/${entitySet}(${id})`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok && res.status !== 404) {
    throw new Error(`DELETE ${entitySet}(${id}): ${res.status}\n${await res.text()}`);
  }
}

async function createRecord(entitySet, payload) {
  const res = await fetch(`${ENV_URL}/api/data/v9.2/${entitySet}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    throw new Error(`POST ${entitySet}: ${res.status}\n${await res.text()}`);
  }
  return await res.json();
}

// ── 1. Resolve prompt IDs by promptkey ──────────────────────────────
console.log("=== Resolving prompt IDs by promptkey ===");
const prompts = await listAll("csp_aiprompts", "csp_aipromptid,csp_promptkey,csp_name");
const promptByKey = Object.fromEntries(
  prompts.map((p) => [p.csp_promptkey, { id: p.csp_aipromptid, name: p.csp_name }])
);
for (const k of ["briefing", "weekly-blog", "weekly-insight"]) {
  if (!promptByKey[k]) {
    throw new Error(`Required prompt not found: ${k}. Run seed-prompts.mjs first.`);
  }
  console.log(`  ✓ ${k} → ${promptByKey[k].id}`);
}
console.log("");

// ── 2. Wipe existing configs (idempotent re-seed) ────────────────────
console.log("=== Deleting existing summary configs ===");
const existing = await listAll("csp_aisummaryconfigs", "csp_aisummaryconfigid,csp_name");
console.log(`  ${existing.length} config(s) to delete`);
for (const c of existing) {
  await deleteRecord("csp_aisummaryconfigs", c.csp_aisummaryconfigid);
}
console.log("  done.\n");

// ── 3. Define the 3 example configs ──────────────────────────────────

const CONFIGS = [
  {
    name: "Top Accounts · Daily Briefings",
    description:
      "Generates a per-account briefing every weekday morning for the 5 most recently active accounts. Demonstrates PerRecord mode — the prompt runs once per record returned by the FetchXML.",
    cadence: CADENCE.Daily,
    mode: MODE.PerRecord,
    promptKey: "briefing",
    targetEntity: "account",
    queryMode: QUERY_MODE.CustomFetchXML,
    fetchXml: `<fetch top="5">
  <entity name="account">
    <attribute name="accountid" />
    <attribute name="name" />
    <attribute name="address1_city" />
    <attribute name="address1_country" />
    <attribute name="revenue" />
    <attribute name="numberofemployees" />
    <filter type="and">
      <condition attribute="modifiedon" operator="last-x-days" value="30" />
    </filter>
    <order attribute="modifiedon" descending="true" />
  </entity>
</fetch>`,
    relatedFetchXml: `<fetch top="8">
  <entity name="csp_contactlogentry">
    <attribute name="csp_contactlogentryid" />
    <attribute name="csp_summary" />
    <attribute name="csp_sentiment" />
    <attribute name="csp_interactiontype" />
    <attribute name="csp_interactiondate" />
    <filter type="and">
      <condition attribute="csp_account" operator="eq" value="{{ recordId }}" />
    </filter>
    <order attribute="csp_interactiondate" descending="true" />
  </entity>
</fetch>`,
    maxRecords: 5,
    language: "en-US",
    outputDestination: OUTPUT_DESTINATION.Cache,
  },
  {
    name: "Weekly Activity Recap",
    description:
      "Aggregates every interaction logged in the past 7 days into a single weekly debrief blog. Demonstrates Aggregate mode — the prompt runs once with all records as a single context payload.",
    cadence: CADENCE.Weekly,
    mode: MODE.Aggregate,
    promptKey: "weekly-blog",
    targetEntity: "csp_contactlogentry",
    queryMode: QUERY_MODE.CustomFetchXML,
    fetchXml: `<fetch top="100">
  <entity name="csp_contactlogentry">
    <attribute name="csp_contactlogentryid" />
    <attribute name="csp_name" />
    <attribute name="csp_summary" />
    <attribute name="csp_sentiment" />
    <attribute name="csp_interactiontype" />
    <attribute name="csp_interactiondate" />
    <attribute name="csp_account" />
    <attribute name="csp_proposednextaction" />
    <filter type="and">
      <condition attribute="csp_interactiondate" operator="last-x-days" value="7" />
    </filter>
    <order attribute="csp_interactiondate" descending="true" />
  </entity>
</fetch>`,
    maxRecords: 100,
    language: "en-US",
    outputDestination: OUTPUT_DESTINATION.Cache,
  },
  {
    name: "Weekly Risk Insight",
    description:
      "Surfaces one sharp insight from the week's risky interactions only. Same engine as the weekly recap but a tighter query (only Risk/Negative sentiment) and a different prompt (weekly-insight). Shows that the same FetchXML pattern feeds completely different prompts.",
    cadence: CADENCE.Weekly,
    mode: MODE.Aggregate,
    promptKey: "weekly-insight",
    targetEntity: "csp_contactlogentry",
    queryMode: QUERY_MODE.CustomFetchXML,
    fetchXml: `<fetch top="50">
  <entity name="csp_contactlogentry">
    <attribute name="csp_contactlogentryid" />
    <attribute name="csp_summary" />
    <attribute name="csp_sentiment" />
    <attribute name="csp_competitorsmentioned" />
    <attribute name="csp_interactiondate" />
    <attribute name="csp_account" />
    <filter type="and">
      <condition attribute="csp_interactiondate" operator="last-x-days" value="7" />
      <filter type="or">
        <condition attribute="csp_sentiment" operator="eq" value="100000002" />
        <condition attribute="csp_sentiment" operator="eq" value="100000003" />
      </filter>
    </filter>
    <order attribute="csp_interactiondate" descending="true" />
  </entity>
</fetch>`,
    maxRecords: 50,
    language: "en-US",
    outputDestination: OUTPUT_DESTINATION.Cache,
  },
];

// ── 4. Create the configs ────────────────────────────────────────────
console.log("=== Creating 3 example summary configs ===");
for (const c of CONFIGS) {
  const promptId = promptByKey[c.promptKey].id;
  const payload = {
    csp_name: c.name,
    csp_description: c.description,
    csp_cadence: c.cadence,
    csp_mode: c.mode,
    "csp_Prompt@odata.bind": `/csp_aiprompts(${promptId})`,
    csp_targetentity: c.targetEntity,
    csp_querymode: c.queryMode,
    csp_fetchxml: c.fetchXml,
    csp_maxrecords: c.maxRecords,
    csp_language: c.language,
    csp_outputdestination: c.outputDestination,
    csp_notifyrecipients: JSON.stringify({ users: [], teams: [], emails: [] }),
    csp_enabled: true,
  };
  if (c.relatedFetchXml) {
    payload.csp_relatedfetchxml = c.relatedFetchXml;
  }
  await createRecord("csp_aisummaryconfigs", payload);
  console.log(`  ✓ ${c.name} (${c.promptKey} · mode=${c.mode === MODE.PerRecord ? "PerRecord" : "Aggregate"})`);
  await sleep(80);
}

console.log("\nDone. 3 summary configs seeded.\n");
