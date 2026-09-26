/**
 * Seed synthetic data for the Irish Power Platform Summit demo.
 *
 * Context: SMBC Aviation Capital (smbc.aero) — one of the world's largest
 * aircraft lessors. The "sales reps" in this product are SMBC Aviation
 * Capital account managers; "accounts" are the airlines they sell to
 * (lease aircraft to). All interactions logged are between an SMBC account
 * manager and airline procurement/fleet/finance contacts.
 *
 * Every account and contact created here is flagged csp_isdemo = true so
 * clean-demo-data.mjs can wipe them later without touching real data.
 * Custom tables (csp_*) are populated only with demo data, so they get
 * truncated wholesale on cleanup.
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

// ─── airline customers ─────────────────────────────────────────────────────

const AIRLINES = [
  {
    name: "Lufthansa Group",
    industryCode: 12, // Transportation
    revenue: 35000000000,
    employees: 105000,
    address: { city: "Frankfurt", country: "Germany", street: "Lufthansa Aviation Center" },
    website: "https://www.lufthansagroup.com",
    health: { score: 84, band: 100000000 /* Green */, signals: ["Strong renewal pipeline", "Recent A350 order signed", "On-time payment 24m"] },
    contacts: [
      { firstName: "Klaus", lastName: "Bauer", title: "Group CFO", email: "klaus.bauer@lufthansa-demo.com" },
      { firstName: "Anke", lastName: "Müller", title: "VP Fleet Management", email: "anke.muller@lufthansa-demo.com" },
      { firstName: "Stefan", lastName: "Hoffmann", title: "Head of Aircraft Procurement", email: "stefan.hoffmann@lufthansa-demo.com" }
    ]
  },
  {
    name: "International Airlines Group (IAG)",
    industryCode: 12,
    revenue: 32000000000,
    employees: 73000,
    address: { city: "Madrid", country: "Spain", street: "Calle de Velázquez 130" },
    website: "https://www.iairgroup.com",
    health: { score: 78, band: 100000000 /* Green */, signals: ["Fleet expansion in progress", "Profitable last 4 quarters"] },
    contacts: [
      { firstName: "Marta", lastName: "Sánchez", title: "Head of Fleet Strategy", email: "marta.sanchez@iag-demo.com" },
      { firstName: "Carlos", lastName: "Ramírez", title: "CFO Iberia", email: "carlos.ramirez@iag-demo.com" }
    ]
  },
  {
    name: "Ryanair Holdings",
    industryCode: 12,
    revenue: 13500000000,
    employees: 22000,
    address: { city: "Dublin", country: "Ireland", street: "Airside Business Park, Swords" },
    website: "https://corporate.ryanair.com",
    health: { score: 62, band: 100000001 /* Yellow */, signals: ["Hard negotiation on lease rates", "Pushing for shorter terms", "Competitor evaluating"] },
    contacts: [
      { firstName: "Liam", lastName: "O'Brien", title: "Director of Fleet", email: "liam.obrien@ryanair-demo.com" },
      { firstName: "Siobhán", lastName: "Murphy", title: "Senior Procurement Manager", email: "siobhan.murphy@ryanair-demo.com" }
    ]
  },
  {
    name: "Wizz Air Holdings",
    industryCode: 12,
    revenue: 5200000000,
    employees: 7000,
    address: { city: "Budapest", country: "Hungary", street: "Kőér utca 2/A" },
    website: "https://wizzair.com",
    health: { score: 71, band: 100000001 /* Yellow */, signals: ["A321neo deliveries delayed", "FX exposure raised by CFO"] },
    contacts: [
      { firstName: "József", lastName: "Váradi", title: "CEO Office Liaison", email: "jvaradi@wizz-demo.com" },
      { firstName: "Erika", lastName: "Kovács", title: "Fleet Planning Lead", email: "erika.kovacs@wizz-demo.com" }
    ]
  },
  {
    name: "easyJet plc",
    industryCode: 12,
    revenue: 9000000000,
    employees: 14000,
    address: { city: "Luton", country: "United Kingdom", street: "Hangar 89, London Luton Airport" },
    website: "https://corporate.easyjet.com",
    health: { score: 75, band: 100000000 /* Green */, signals: ["Confirmed 5 A320neo extension"] },
    contacts: [
      { firstName: "James", lastName: "Whitmore", title: "Director Aircraft Trading", email: "james.whitmore@easyjet-demo.com" },
      { firstName: "Priya", lastName: "Patel", title: "Lease Manager", email: "priya.patel@easyjet-demo.com" }
    ]
  },
  {
    name: "Aeroméxico",
    industryCode: 12,
    revenue: 4100000000,
    employees: 14500,
    address: { city: "Mexico City", country: "Mexico", street: "Av. Paseo de la Reforma 243" },
    website: "https://www.aeromexico.com",
    health: { score: 45, band: 100000002 /* Red */, signals: ["Late payment x2 last quarter", "Talks with KKR Aviation", "Contract expires in 67 days"] },
    contacts: [
      { firstName: "Ricardo", lastName: "Mendoza", title: "Chief Financial Officer", email: "rmendoza@aeromexico-demo.com" },
      { firstName: "Lucía", lastName: "Hernández", title: "VP Fleet & Network", email: "lhernandez@aeromexico-demo.com" }
    ]
  },
  {
    name: "All Nippon Airways (ANA)",
    industryCode: 12,
    revenue: 17500000000,
    employees: 46500,
    address: { city: "Tokyo", country: "Japan", street: "Shiodome City Center, Minato" },
    website: "https://www.ana.co.jp",
    health: { score: 88, band: 100000000 /* Green */, signals: ["Long-standing partner", "B787 fleet expansion confirmed", "NPS 9.2"] },
    contacts: [
      { firstName: "Hiroshi", lastName: "Tanaka", title: "Senior VP Fleet Strategy", email: "hiroshi.tanaka@ana-demo.com" },
      { firstName: "Yuki", lastName: "Sato", title: "Procurement Director", email: "yuki.sato@ana-demo.com" }
    ]
  },
  {
    name: "LATAM Airlines Group",
    industryCode: 12,
    revenue: 12000000000,
    employees: 33000,
    address: { city: "Santiago", country: "Chile", street: "Avenida Presidente Riesco 5711" },
    website: "https://www.latamairlinesgroup.net",
    health: { score: 68, band: 100000001 /* Yellow */, signals: ["Recovering post-Chapter 11", "Cautious on new commitments"] },
    contacts: [
      { firstName: "Ana", lastName: "Vergara", title: "Director Fleet Planning", email: "ana.vergara@latam-demo.com" },
      { firstName: "Diego", lastName: "Rojas", title: "Treasury Lead", email: "diego.rojas@latam-demo.com" }
    ]
  }
];

// ─── interaction templates ─────────────────────────────────────────────────

const INTERACTION_TEMPLATES = [
  {
    type: 100000000, // Call
    sentiment: 100000002, // Negative
    raw: "Just got off the phone with {contact}. He's pushing back hard on the lease rate for the 3 A320neo extension — says KKR Aviation quoted 8% less. Wants us to revise by end of week or they walk.",
    summary: "Pricing pressure on A320neo extension. Competitor KKR Aviation quoted 8% lower. Decision deadline end of week.",
    competitors: "KKR Aviation",
    products: "A320neo",
    nextAction: "Revise lease rate proposal before Friday; consider shorter term to bridge the gap"
  },
  {
    type: 100000002, // Meeting
    sentiment: 100000000, // Positive
    raw: "Met with {contact} in {city}. Confirmed acceptance of the 5-aircraft remarketing package. Signing expected mid-month. Mentioned interest in our A350-1000 availability for 2027 deliveries.",
    summary: "{contact} accepted the 5-aircraft remarketing package. Signing imminent. Pipeline: A350-1000 for 2027.",
    competitors: "",
    products: "A320neo, A350-1000",
    nextAction: "Send signed term sheet by Wednesday and start A350 availability discussion next month"
  },
  {
    type: 100000001, // Email
    sentiment: 100000001, // Neutral
    raw: "{contact} replied to our QBR follow-up. Confirmed numbers from last quarter look good but flagged that their CFO wants to see total cost of ownership analysis including maintenance reserves before the next renewal.",
    summary: "QBR follow-up: CFO wants TCO analysis including maintenance reserves before renewal discussion proceeds.",
    competitors: "",
    products: "",
    nextAction: "Prepare TCO model with current MR rates and share by next Tuesday"
  },
  {
    type: 100000000, // Call
    sentiment: 100000003, // Risk
    raw: "Urgent call from {contact}. Their A330neo deliveries are slipping by 4 months due to engine OEM issues. They are exploring alternative wide-body lessors. Risk of losing the renewal.",
    summary: "Risk: A330neo delivery delays from OEM. Customer considering alternative wide-body lessors. Renewal at risk.",
    competitors: "Avolon, BBAM",
    products: "A330neo",
    nextAction: "Escalate internally to discuss compensation package; schedule call with CEO of customer within 48h"
  },
  {
    type: 100000003, // Note
    sentiment: 100000001, // Neutral
    raw: "Quarterly check-in. {contact} mentioned they're starting to evaluate sustainable aviation fuel (SAF) integration with new aircraft deliveries. Possible angle for our 2027 sustainability proposition.",
    summary: "Customer interested in SAF integration with new deliveries. Potential angle for sustainability-linked lease terms.",
    competitors: "",
    products: "",
    nextAction: "Brief our sustainability team and prepare SAF-linked lease pitch for Q3"
  },
  {
    type: 100000002, // Meeting
    sentiment: 100000000, // Positive
    raw: "Site visit at {city} HQ. {contact} gave us a tour of the new operations center. Very impressed with our remarketing process. They want to discuss expanding the partnership beyond pure leasing — possibly tech support services.",
    summary: "Site visit went well. Customer interested in expanding partnership to include technical support services.",
    competitors: "",
    products: "",
    nextAction: "Coordinate with our tech support BU to scope an offering and present in 3 weeks"
  },
  {
    type: 100000000, // Call
    sentiment: 100000002, // Negative
    raw: "{contact} called frustrated about return conditions on the 2 A321 coming off lease in Q3. Says our reps were too strict in pre-redelivery inspection. Threatened to escalate to CEO if not resolved.",
    summary: "Conflict on return condition inspection for 2 A321. Customer feels overly strict assessment. Escalation risk.",
    competitors: "",
    products: "A321",
    nextAction: "Send technical lead to do joint walkaround inspection this week; document agreed deltas"
  },
  {
    type: 100000001, // Email
    sentiment: 100000000, // Positive
    raw: "{contact} confirmed via email that the maintenance reserve rate adjustment we proposed last quarter is approved. Will start applying from January. Asked us to include the language in the master agreement amendment.",
    summary: "MR rate adjustment approved. Will apply from January. Customer requests master agreement amendment.",
    competitors: "",
    products: "",
    nextAction: "Have legal prepare master agreement amendment and send for signature within 10 days"
  },
  {
    type: 100000003, // Note
    sentiment: 100000003, // Risk
    raw: "Heard from industry contact that {contact}'s company is in early-stage discussions with a competitor for a portfolio swap. Not confirmed but worth tracking. Their lease portfolio with us is 14 aircraft.",
    summary: "Unconfirmed market intel: customer may be discussing portfolio swap with competitor. 14 aircraft at stake.",
    competitors: "AerCap",
    products: "Portfolio (14 a/c)",
    nextAction: "Schedule informal CEO-to-CEO call to test the temperature; do not raise the rumor directly"
  },
  {
    type: 100000002, // Meeting
    sentiment: 100000000, // Positive
    raw: "Pitch meeting with {contact} for 2 new A220-300 in 2027. They liked the proposal. Decision expected in 4-6 weeks pending board approval. Asked for references from other A220 operators in our portfolio.",
    summary: "A220-300 pitch went well. Board decision expected in 4-6 weeks. Customer requests reference list.",
    competitors: "ALC, Aircastle",
    products: "A220-300",
    nextAction: "Prepare A220 operator reference document with permission; share next week"
  },
  {
    type: 100000000, // Call
    sentiment: 100000001, // Neutral
    raw: "Standard quarterly call with {contact}. Numbers tracking to plan. No major issues. Briefly discussed industry trends — pilot shortage, fuel hedging. Nothing actionable but useful color.",
    summary: "Routine quarterly call. No issues. General industry context: pilot shortage, fuel hedging.",
    competitors: "",
    products: "",
    nextAction: "Continue regular cadence; next check-in in 90 days"
  },
  {
    type: 100000001, // Email
    sentiment: 100000003, // Risk
    raw: "{contact} flagged via email that their parent group is going through restructuring. The aviation division may be carved out or spun off. Will affect all our contracts in 12-18 months. Asked to keep this confidential.",
    summary: "Confidential: customer parent group restructuring. Aviation division may be carved out within 12-18 months. Affects all contracts.",
    competitors: "",
    products: "",
    nextAction: "Brief our credit team; assess exposure across all contracts with this customer"
  }
];

// ─── next best action templates ────────────────────────────────────────────

const NBA_TEMPLATES = [
  { actionType: 100000000 /* Email */, priority: 100000000 /* High */, description: "Send revised remarketing proposal for the A320neo extension package.", justification: "Customer explicitly asked for revised pricing before Friday. Competitor (KKR Aviation) quoted 8% less and customer threatened to walk if no response." },
  { actionType: 100000003 /* Meeting */, priority: 100000000 /* High */, description: "Schedule QBR meeting with customer CFO and Fleet Director.", justification: "Master agreement contract expires in 67 days and no renewal conversation has been initiated. Health score is in the Red band and customer is in active talks with competitor." },
  { actionType: 100000001 /* Call */, priority: 100000001 /* Medium */, description: "Call procurement lead to discuss A350-1000 availability for 2027.", justification: "In last meeting customer expressed interest in A350-1000. Our pipeline has 4 available slots for 2027 deliveries and pricing competitive vs ALC and BBAM." },
  { actionType: 100000002 /* Task */, priority: 100000001 /* Medium */, description: "Prepare TCO model with current maintenance reserve rates.", justification: "CFO explicitly requested a Total Cost of Ownership analysis before approving the renewal. Internal Finance can produce in 5 business days." },
  { actionType: 100000004 /* Proposal */, priority: 100000001 /* Medium */, description: "Draft SAF-linked lease term sheet as discussed in last visit.", justification: "Customer surfaced interest in sustainable aviation fuel integration. Sustainability-linked terms are a strong differentiator vs traditional lessors and align with their 2030 net-zero commitment." },
  { actionType: 100000001 /* Call */, priority: 100000000 /* High */, description: "Escalate A330neo OEM delay to senior leadership for compensation package.", justification: "Customer at risk of walking to alternative wide-body lessor over delivery delays we can't fully control. Internal escalation needed to authorize a compensation envelope to retain the renewal." },
  { actionType: 100000003 /* Meeting */, priority: 100000002 /* Low */, description: "Joint inspection walkaround for the 2 A321 returning in Q3.", justification: "Customer raised concern about strictness of pre-redelivery inspection. A joint walkaround would defuse the conflict and reduce escalation risk to CEO level." },
  { actionType: 100000002 /* Task */, priority: 100000001 /* Medium */, description: "Get permission from 2 reference A220 operators in our portfolio and share with customer.", justification: "Customer requested references as part of the A220-300 pitch evaluation. Board decision expected in 4-6 weeks." }
];

// ─── auth and api helpers ──────────────────────────────────────────────────

const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`
  }
});

async function getToken() {
  const r = await pca.acquireTokenByDeviceCode({
    scopes: [`${ENV_URL}/.default`],
    deviceCodeCallback: (resp) => {
      console.log("\n=========================================");
      console.log("DEVICE CODE AUTHENTICATION");
      console.log("=========================================");
      console.log(resp.message);
      console.log("=========================================\n");
    }
  });
  return r.accessToken;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let token;

async function api(method, urlPath, body) {
  const res = await fetch(`${ENV_URL}/api/data/v9.2${urlPath}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Prefer: "return=representation"
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${method} ${urlPath} -> ${res.status}\n${text}`);
  }
  if (res.status === 204) {
    const id = res.headers.get("OData-EntityId");
    return id ? { id: id.match(/\(([^)]+)\)/)?.[1] } : null;
  }
  return res.json();
}

async function createRecord(entitySet, data) {
  const result = await api("POST", `/${entitySet}`, data);
  await sleep(150);
  return result;
}

// ─── helpers ───────────────────────────────────────────────────────────────

function pick(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

function pickN(arr, n) {
  const shuffled = [...arr].sort(() => Math.random() - 0.5);
  return shuffled.slice(0, Math.min(n, arr.length));
}

function randInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function daysAgo(d) {
  const date = new Date();
  date.setDate(date.getDate() - d);
  return date.toISOString();
}

function getISOWeek(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  const week1 = new Date(d.getFullYear(), 0, 4);
  return 1 + Math.round(((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
}

// ─── main ──────────────────────────────────────────────────────────────────

// resolve navigation property names for lookup fields at runtime
async function resolveNavProps(entityLogical) {
  const data = await api(
    "GET",
    `/EntityDefinitions(LogicalName='${entityLogical}')?$select=LogicalName&$expand=ManyToOneRelationships($select=ReferencingAttribute,ReferencingEntityNavigationPropertyName,ReferencedEntity)`
  );
  const map = {};
  for (const rel of data.ManyToOneRelationships || []) {
    map[rel.ReferencingAttribute] = rel.ReferencingEntityNavigationPropertyName;
  }
  return map;
}

async function main() {
  console.log("Authenticating against tenant...");
  token = await getToken();
  console.log("Authenticated.\n");

  // current user (sales rep / weekly recap owner)
  const me = await api("GET", "/WhoAmI()");
  console.log(`Sales rep (current user): ${me.UserId}\n`);

  // discover navigation property names (different from lookup logical name)
  console.log("Resolving navigation properties for custom tables...");
  const nav = {
    csp_contactlogentry: await resolveNavProps("csp_contactlogentry"),
    csp_nextbestaction: await resolveNavProps("csp_nextbestaction"),
    csp_accounthealth: await resolveNavProps("csp_accounthealth"),
    csp_aisummarycache: await resolveNavProps("csp_aisummarycache"),
    csp_aiusage: await resolveNavProps("csp_aiusage")
  };
  console.log("Resolved.\n");
  for (const [t, m] of Object.entries(nav)) {
    for (const [attr, prop] of Object.entries(m)) {
      console.log(`  ${t}.${attr} → ${prop}`);
    }
  }
  console.log();

  // ─── 1. Accounts (airlines) + contacts ───────────────────────────────────
  console.log("=== Creating accounts and contacts ===");
  const createdAccounts = [];

  for (const airline of AIRLINES) {
    console.log(`  [account] ${airline.name}`);
    const account = await createRecord("accounts", {
      name: airline.name,
      industrycode: airline.industryCode,
      revenue: airline.revenue,
      numberofemployees: airline.employees,
      address1_city: airline.address.city,
      address1_country: airline.address.country,
      address1_line1: airline.address.street,
      websiteurl: airline.website,
      description: `Synthetic demo account representing ${airline.name} as a hypothetical airline customer of SMBC Aviation Capital. Not real customer data.`,
      csp_isdemo: true
    });

    const contacts = [];
    for (const c of airline.contacts) {
      console.log(`    [contact] ${c.firstName} ${c.lastName} (${c.title})`);
      const contact = await createRecord("contacts", {
        firstname: c.firstName,
        lastname: c.lastName,
        jobtitle: c.title,
        emailaddress1: c.email,
        "parentcustomerid_account@odata.bind": `/accounts(${account.accountid})`,
        csp_isdemo: true
      });
      contacts.push(contact);
    }

    createdAccounts.push({ airline, accountId: account.accountid, contacts });
  }

  // ─── 2. Contact log entries (interactions) ──────────────────────────────
  console.log("\n=== Creating contact log entries ===");
  let totalInteractions = 0;
  for (const acc of createdAccounts) {
    const numInteractions = randInt(2, 4);
    const templates = pickN(INTERACTION_TEMPLATES, numInteractions);
    for (let i = 0; i < templates.length; i++) {
      const tpl = templates[i];
      const contact = pick(acc.contacts);
      const contactName = `${acc.airline.contacts.find((c) => c.email === contact.emailaddress1)?.firstName ?? "the contact"}`;
      const dayOffset = randInt(1, 90);
      const interaction = {
        csp_name: `${acc.airline.name} - ${tpl.summary.substring(0, 80)}`,
        csp_rawinput: tpl.raw.replaceAll("{contact}", contactName).replaceAll("{city}", acc.airline.address.city),
        csp_summary: tpl.summary.replaceAll("{contact}", contactName),
        csp_sentiment: tpl.sentiment,
        csp_interactiontype: tpl.type,
        csp_competitorsmentioned: tpl.competitors,
        csp_productsmentioned: tpl.products,
        csp_proposednextaction: tpl.nextAction,
        csp_interactiondate: daysAgo(dayOffset),
        [`${nav.csp_contactlogentry.csp_account}@odata.bind`]: `/accounts(${acc.accountId})`,
        [`${nav.csp_contactlogentry.csp_contact}@odata.bind`]: `/contacts(${contact.contactid})`
      };
      await createRecord("csp_contactlogentries", interaction);
      totalInteractions++;
    }
    console.log(`  [interactions] ${acc.airline.name}: ${numInteractions} logged`);
  }
  console.log(`  Total interactions: ${totalInteractions}`);

  // ─── 3. Next Best Actions ───────────────────────────────────────────────
  console.log("\n=== Creating next best actions ===");
  let totalNBAs = 0;
  for (const acc of createdAccounts) {
    const numNBAs = randInt(1, 2);
    const templates = pickN(NBA_TEMPLATES, numNBAs);
    for (const tpl of templates) {
      await createRecord("csp_nextbestactions", {
        csp_name: `${acc.airline.name} - ${tpl.description.substring(0, 80)}`,
        csp_description: tpl.description,
        csp_justification: tpl.justification,
        csp_actiontype: tpl.actionType,
        csp_priority: tpl.priority,
        csp_status: 100000000, // Pending
        csp_generatedon: daysAgo(randInt(0, 5)),
        [`${nav.csp_nextbestaction.csp_account}@odata.bind`]: `/accounts(${acc.accountId})`
      });
      totalNBAs++;
    }
  }
  console.log(`  Total NBAs: ${totalNBAs}`);

  // ─── 4. Account Health snapshots ────────────────────────────────────────
  console.log("\n=== Creating account health snapshots ===");
  for (const acc of createdAccounts) {
    const h = acc.airline.health;
    // current snapshot
    await createRecord("csp_accounthealths", {
      csp_name: `${acc.airline.name} - Health ${new Date().toISOString().substring(0, 10)}`,
      csp_score: h.score,
      csp_band: h.band,
      csp_signals: JSON.stringify(h.signals),
      csp_calculatedon: new Date().toISOString(),
      [`${nav.csp_accounthealth.csp_account}@odata.bind`]: `/accounts(${acc.accountId})`
    });
    // 2 weeks of history for trend
    for (let w = 1; w <= 2; w++) {
      const drift = randInt(-7, 7);
      const histScore = Math.max(0, Math.min(100, h.score + drift));
      await createRecord("csp_accounthealths", {
        csp_name: `${acc.airline.name} - Health W-${w}`,
        csp_score: histScore,
        csp_band: histScore >= 75 ? 100000000 : histScore >= 55 ? 100000001 : 100000002,
        csp_signals: JSON.stringify(["Historical snapshot"]),
        csp_calculatedon: daysAgo(w * 7),
        [`${nav.csp_accounthealth.csp_account}@odata.bind`]: `/accounts(${acc.accountId})`
      });
    }
  }
  console.log(`  Health snapshots: ${createdAccounts.length * 3}`);

  // ─── 5. AI Summary Cache (briefing per account) ─────────────────────────
  console.log("\n=== Creating AI summary cache (briefings) ===");
  for (const acc of createdAccounts) {
    const h = acc.airline.health;
    const briefing = `## ${acc.airline.name} - Executive Briefing

**Headquarters**: ${acc.airline.address.city}, ${acc.airline.address.country}
**Revenue**: ${(acc.airline.revenue / 1e9).toFixed(1)}B EUR · **Employees**: ${acc.airline.employees.toLocaleString()}

### Current health
Score **${h.score}/100** (${h.band === 100000000 ? "Green" : h.band === 100000001 ? "Yellow" : "Red"}).
${h.signals.map((s) => `- ${s}`).join("\n")}

### Relationship snapshot
${acc.airline.name} is a long-standing customer of SMBC Aviation Capital. Current portfolio includes leased single-aisle and wide-body aircraft. Key contacts in fleet, procurement and finance functions.

### Suggested focus
Review open Next Best Actions and prioritize ${h.band === 100000002 ? "retention" : h.band === 100000001 ? "stabilization" : "expansion"} initiatives this quarter.`;

    const tokensUsed = randInt(800, 1800);
    await createRecord("csp_aisummarycaches", {
      csp_name: `${acc.airline.name} - Briefing`,
      csp_summarytype: 100000000, // Briefing
      csp_content: briefing,
      csp_generatedon: daysAgo(randInt(0, 3)),
      csp_tokensused: tokensUsed,
      csp_promptversion: "v1.0",
      [`${nav.csp_aisummarycache.csp_account}@odata.bind`]: `/accounts(${acc.accountId})`
    });
  }
  console.log(`  Briefings cached: ${createdAccounts.length}`);

  // ─── 6. Weekly Recaps (last 4 weeks for current user) ───────────────────
  console.log("\n=== Creating weekly recaps ===");
  const now = new Date();
  for (let w = 0; w < 4; w++) {
    const weekStart = new Date(now);
    weekStart.setDate(weekStart.getDate() - w * 7);
    const weekNum = getISOWeek(weekStart);
    const year = weekStart.getFullYear();
    const blogTitle = pick([
      "A week of pivots: pricing pressure mounts, two renewals signed",
      "Quiet week on the headline deals, busy week on the credit risk side",
      "Two wins, one risk: portfolio holds steady",
      "From Madrid to Tokyo: cross-region momentum building"
    ]);
    const blogContent = `## ${blogTitle}

This week was defined by **pricing pressure on the single-aisle book** and **two confirmed signings on the wide-body side**.

### Most important
- Ryanair pushed back hard on the A320neo extension; we have until Friday to revise terms.
- IAG signed the 5-aircraft remarketing package — finally over the line after 6 weeks of negotiation.
- Aeroméxico moved into Red band: late payment x2 last quarter plus active competitor talks.

### Accounts that advanced
- IAG — 5 aircraft package signed.
- ANA — A350-1000 pitch went well, board decision in 4-6 weeks.

### Accounts at risk
- Ryanair — competitor pricing 8% lower; revise or walk.
- Aeroméxico — Red health, payment delays, KKR Aviation in talks.

### Next week's agenda
Six Next Best Actions are pending. Top priorities: revised Ryanair proposal, Aeroméxico QBR with CFO, and TCO model for Lufthansa renewal.

### Insight of the week
Three of your top six pending NBAs are tied to **competitor pricing pressure on single-aisle**. That's a pattern, not a coincidence — worth raising at the next sales leadership meeting.`;

    await createRecord("csp_weeklyrecaps", {
      csp_name: `Week ${weekNum} - ${year}`,
      csp_weeknumber: weekNum,
      csp_year: year,
      csp_blogtitle: blogTitle,
      csp_blogcontent: blogContent,
      csp_audiourl: `https://example.blob.core.windows.net/podcasts/week-${year}-${weekNum}.mp3`,
      csp_audioduration: randInt(240, 420),
      csp_coverurl: `https://example.blob.core.windows.net/covers/week-${year}-${weekNum}.png`,
      csp_highlights: JSON.stringify([
        "IAG signed 5-aircraft package",
        "Ryanair pricing pressure",
        "Aeroméxico into Red band"
      ]),
      csp_insight: "Three of six pending NBAs are tied to competitor pricing pressure on single-aisle.",
      csp_tokensused: randInt(3000, 5000)
    });
  }
  console.log(`  Weekly recaps: 4`);

  // ─── 7. AI Usage logs ───────────────────────────────────────────────────
  console.log("\n=== Creating AI usage logs ===");
  const promptNames = ["prompt_account_briefing", "prompt_contactlog_extraction", "prompt_next_best_actions", "prompt_weekly_blog", "prompt_weekly_insight"];
  const models = ["gpt-4o", "gpt-4o-mini"];
  let usageCount = 0;
  for (let d = 0; d < 14; d++) {
    const callsToday = randInt(3, 7);
    for (let i = 0; i < callsToday; i++) {
      const prompt = pick(promptNames);
      const model = pick(models);
      const tokensIn = randInt(500, 4000);
      const tokensOut = randInt(200, 1500);
      const costPer1K = model === "gpt-4o" ? 0.0025 : 0.00015;
      const cost = ((tokensIn + tokensOut) / 1000) * costPer1K;
      const acc = pick(createdAccounts);
      await createRecord("csp_aiusages", {
        csp_name: `${prompt} - ${daysAgo(d).substring(0, 10)} #${i}`,
        csp_promptname: prompt,
        csp_tokensinput: tokensIn,
        csp_tokensoutput: tokensOut,
        csp_estimatedcost: parseFloat(cost.toFixed(4)),
        csp_latencyms: randInt(800, 4500),
        csp_model: model,
        csp_timestamp: daysAgo(d),
        [`${nav.csp_aiusage.csp_account}@odata.bind`]: `/accounts(${acc.accountId})`,
        [`${nav.csp_aiusage.csp_user}@odata.bind`]: `/systemusers(${me.UserId})`
      });
      usageCount++;
    }
  }
  console.log(`  AI usage entries: ${usageCount}`);

  console.log("\n=== Summary ===");
  console.log(`Accounts: ${createdAccounts.length}`);
  console.log(`Contacts: ${createdAccounts.reduce((s, a) => s + a.contacts.length, 0)}`);
  console.log(`Interactions: ${totalInteractions}`);
  console.log(`NBAs: ${totalNBAs}`);
  console.log(`Health snapshots: ${createdAccounts.length * 5}`);
  console.log(`Briefings: ${createdAccounts.length}`);
  console.log(`Weekly recaps: 4`);
  console.log(`AI usage entries: ${usageCount}`);
  console.log("\nDone.");
}

main().catch((err) => {
  console.error("\nFATAL:", err.message);
  process.exit(1);
});
