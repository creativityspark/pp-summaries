/**
 * Seed the 5 baseline AI prompts into csp_aiprompts.
 *
 * Each prompt corresponds to one of the AI flows that power the
 * Sales Briefing Studio. They are editable end-to-end from the
 * AI Configuration page in the app once seeded.
 *
 * Prompts target an Azure AI Foundry deployment (see DEFAULT_DEPLOYMENT
 * below). The Power Automate flow renders the Liquid csp_content
 * server-side via the custom Liquid Render API, then posts the rendered
 * prompt to the Foundry endpoint via the AI Foundry custom connector.
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

// Confirmed defaults for the Irish Power Platform Summit 2026 demo.
const DEFAULT_MODEL = "gpt-4o-mini";
const DEFAULT_DEPLOYMENT = "gpt-4o-mini";

// csp_outputformat picklist values — keep in sync with tables.json.
const OUTPUT_FORMAT = {
  Markdown: 100000000,
  Json: 100000001,
  Html: 100000002,
  Text: 100000003,
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

async function listIds(entitySet, primaryKey) {
  const all = [];
  let url = `${ENV_URL}/api/data/v9.2/${entitySet}?$select=${primaryKey}`;
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
    for (const r of data.value) all.push(r[primaryKey]);
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

const now = new Date();
const daysAgo = (n) => {
  const d = new Date(now);
  d.setDate(d.getDate() - n);
  return d.toISOString();
};

const EXTRACT_OUTPUT_SCHEMA = {
  type: "object",
  required: [
    "summary",
    "interactionType",
    "sentiment",
    "competitorsMentioned",
    "productsMentioned",
    "proposedNextAction",
  ],
  properties: {
    summary: { type: "string" },
    interactionType: { type: "string", enum: ["Call", "Email", "Meeting", "Note"] },
    sentiment: { type: "string", enum: ["Positive", "Neutral", "Negative", "Risk"] },
    competitorsMentioned: { type: "array", items: { type: "string" } },
    productsMentioned: { type: "array", items: { type: "string" } },
    proposedNextAction: { type: "string" },
  },
  additionalProperties: false,
};

const NBA_OUTPUT_SCHEMA = {
  type: "array",
  minItems: 3,
  maxItems: 5,
  items: {
    type: "object",
    required: ["description", "actionType", "priority", "rationale"],
    properties: {
      description: { type: "string" },
      actionType: {
        type: "string",
        enum: ["Email", "Call", "Task", "Meeting", "Proposal"],
      },
      priority: { type: "string", enum: ["High", "Medium", "Low"] },
      rationale: { type: "string" },
    },
    additionalProperties: false,
  },
};

const PROMPTS = [
  {
    key: "briefing",
    name: "Account Briefing",
    description:
      "Synthesises the latest contact log entries into a one-page account brief.",
    systemMessage:
      "You are an expert account manager assistant for SMBC Aviation Capital. Write in a direct, opinionated voice grounded in specifics. Never pad with generalities.",
    model: DEFAULT_MODEL,
    deploymentName: DEFAULT_DEPLOYMENT,
    version: "v2.0",
    enabled: true,
    temperature: 0.3,
    maxTokens: 1500,
    lastRunDaysAgo: 0,
    tokensLastRun: 1240,
    outputFormat: OUTPUT_FORMAT.Markdown,
    outputSchema: null,
    expectedVariables: [
      {
        name: "account",
        type: "object",
        description:
          "Target account: name, city, country, revenue, healthScore, healthBand, healthSignals[].",
      },
      {
        name: "interactions",
        type: "array",
        description:
          "Recent contact log entries (last 90 days). Each: date, type, sentiment, summary.",
      },
      {
        name: "openOpportunities",
        type: "array",
        description: "Open opportunities. Each: name, value, stage.",
      },
    ],
    content: `Given the account context and the latest interactions, produce a concise one-page briefing that helps the sales rep walk into their next conversation prepared.

{% if account.healthBand == "Red" %}> ⚠️ This account is currently in **Red band** — open the brief with the risk framing, not the relationship history.
{% endif %}
# Account context
- Name: {{ account.name }}
- Location: {{ account.city }}, {{ account.country }}
- Revenue: {{ account.revenue | money }}
- Health: {{ account.healthScore }}/100 ({{ account.healthBand }})
{% if account.healthSignals.size > 0 %}- Active signals:
{% for signal in account.healthSignals %}  - {{ signal }}
{% endfor %}{% endif %}

# Recent interactions (last 90 days)
{% for i in interactions limit:8 %}- [{{ i.date | date: "%d %b" }}] **{{ i.type }}** · {{ i.sentiment }}: {{ i.summary }}
{% endfor %}

# Open opportunities
{% for opp in openOpportunities %}- {{ opp.name }} — {{ opp.value | money }} ({{ opp.stage }})
{% else %}- (none open)
{% endfor %}

# Output format
Return markdown with these sections:

## Account snapshot
2-3 sentence positioning paragraph.

## Current health
One paragraph explaining what's driving the current {{ account.healthBand }} band — cite specific signals, not generalities.

## Relationship snapshot
Who matters on the customer side, what the dynamic is right now, what's working and what isn't.

## Suggested focus
A short, opinionated paragraph telling the rep what to focus on this week — concrete, not generic.

Keep the entire brief under 350 words. Prefer specificity over generality.`,
  },
  {
    key: "extract",
    name: "Contact Log Extraction",
    description:
      "Turns raw call/email/note text into structured fields: sentiment, products, competitors, next action.",
    systemMessage:
      "You extract structured information from sales interaction logs. Respond with valid JSON only. No prose, no markdown, no preamble.",
    model: DEFAULT_MODEL,
    deploymentName: DEFAULT_DEPLOYMENT,
    version: "v2.0",
    enabled: true,
    temperature: 0.1,
    maxTokens: 800,
    lastRunDaysAgo: 0,
    tokensLastRun: 620,
    outputFormat: OUTPUT_FORMAT.Json,
    outputSchema: EXTRACT_OUTPUT_SCHEMA,
    expectedVariables: [
      {
        name: "rawInput",
        type: "string",
        description: "The sales rep's verbatim text describing the interaction.",
      },
      {
        name: "account",
        type: "object",
        description: "Account the interaction belongs to. Required: name.",
      },
      {
        name: "contact",
        type: "object",
        description:
          "Optional contact involved. Properties: fullName, jobTitle. Empty when not known.",
      },
    ],
    content: `Input is the rep's verbatim text describing a call, email, meeting or note with an airline customer.

# Input
{{ rawInput }}

# Account
{{ account.name }}{% if contact %} — primary contact: {{ contact.fullName }}{% if contact.jobTitle %} ({{ contact.jobTitle }}){% endif %}{% endif %}

# Output — JSON only, no prose
Return a single JSON object with these fields:
{
  "summary": "1-2 sentence summary in markdown",
  "interactionType": "Call" | "Email" | "Meeting" | "Note",
  "sentiment": "Positive" | "Neutral" | "Negative" | "Risk",
  "competitorsMentioned": ["competitor1", ...],
  "productsMentioned": ["product1", ...],
  "proposedNextAction": "One concrete next step the rep should take"
}

Guidelines:
- Use "Risk" sentiment only when there is real risk to the relationship or a deal — late payments, competitor active, threat to walk, escalation.
- Keep proposedNextAction concrete and bounded (under 25 words).
- If a field is genuinely empty, return an empty string or empty array — never omit the key.`,
  },
  {
    key: "nba",
    name: "Next Best Actions",
    description:
      "Ranks the top 3-5 actions per account based on health, recent interactions, and open opportunities.",
    systemMessage:
      "You generate prioritised next best actions for an account manager. Respond with a JSON array of 3-5 items only. No prose, no markdown, no preamble.",
    model: DEFAULT_MODEL,
    deploymentName: DEFAULT_DEPLOYMENT,
    version: "v2.0",
    enabled: true,
    temperature: 0.4,
    maxTokens: 1200,
    lastRunDaysAgo: 0,
    tokensLastRun: 980,
    outputFormat: OUTPUT_FORMAT.Json,
    outputSchema: NBA_OUTPUT_SCHEMA,
    expectedVariables: [
      {
        name: "account",
        type: "object",
        description:
          "Target account: name, healthScore, healthBand, healthSignals[].",
      },
      {
        name: "interactions",
        type: "array",
        description:
          "Recent contact log entries. Each: date, type, sentiment, summary.",
      },
      {
        name: "existingNbas",
        type: "array",
        description:
          "Already-pending NBAs to avoid duplicating. Each: actionType, description.",
      },
    ],
    content: `# Account context
- Name: {{ account.name }}
- Health: {{ account.healthScore }}/100 ({{ account.healthBand }})
{% if account.healthSignals.size > 0 %}- Active signals:
{% for signal in account.healthSignals %}  - {{ signal }}
{% endfor %}{% endif %}

# Recent interactions
{% for i in interactions limit:6 %}- [{{ i.date | date: "%d %b" }}] {{ i.type }} · {{ i.sentiment }}: {{ i.summary | truncate: 140 }}
{% endfor %}

# Already-pending actions (do not duplicate)
{% for nba in existingNbas %}- {{ nba.actionType }}: {{ nba.description }}
{% else %}- (none)
{% endfor %}

# Output — JSON array, no prose
Return an array of 3-5 next best actions, each as a JSON object:
[
  {
    "description": "Concrete action sentence (under 30 words)",
    "actionType": "Email" | "Call" | "Task" | "Meeting" | "Proposal",
    "priority": "High" | "Medium" | "Low",
    "rationale": "Why this matters now (under 20 words)"
  }
]

Priority guide:
- High = time-sensitive, related to risk, churn, or deadline within 1 week
- Medium = strategic, important, but no immediate deadline
- Low = nice-to-have, opportunistic

{% if account.healthBand == "Red" %}This account is in **Red band** — at least one High priority action is expected.
{% endif %}Sort the array High → Low. Don't pad with low-value generic actions.`,
  },
  {
    key: "weekly-blog",
    name: "Weekly Recap Blog",
    description:
      "Writes the long-form weekly debrief covering most important moves, wins and risks.",
    systemMessage:
      "You write a weekly sales debrief blog. Your voice is direct, opinionated and grounded in specifics — like a smart colleague catching the rep up over coffee, not a corporate newsletter.",
    model: DEFAULT_MODEL,
    deploymentName: DEFAULT_DEPLOYMENT,
    version: "v2.0",
    enabled: true,
    temperature: 0.5,
    maxTokens: 2500,
    lastRunDaysAgo: 1,
    tokensLastRun: 2180,
    outputFormat: OUTPUT_FORMAT.Markdown,
    outputSchema: null,
    expectedVariables: [
      {
        name: "week",
        type: "object",
        description: "Week context: number, year, dateRange.",
      },
      {
        name: "interactionsThisWeek",
        type: "array",
        description:
          "Interactions logged during the week. Each: accountName, type, sentiment, summary.",
      },
      {
        name: "pendingNbas",
        type: "array",
        description:
          "NBAs still pending into next week. Each: priority, accountName, description.",
      },
      {
        name: "redAccounts",
        type: "array",
        description:
          "Accounts that moved into Red band during the week. Each: name, reason.",
      },
    ],
    content: `# Week
Week {{ week.number }} · {{ week.year }} ({{ week.dateRange }})

# Interactions logged this week ({{ interactionsThisWeek.size }})
{% for i in interactionsThisWeek %}- **{{ i.accountName }}** · {{ i.type }} · {{ i.sentiment }}: {{ i.summary | truncate: 160 }}
{% endfor %}

# Pending NBAs going into next week
{% for nba in pendingNbas %}- {{ nba.priority | upcase }} · **{{ nba.accountName }}** — {{ nba.description }}
{% endfor %}

{% if redAccounts.size > 0 %}# Accounts that moved into Red band this week
{% for acc in redAccounts %}- **{{ acc.name }}** — {{ acc.reason }}
{% endfor %}{% endif %}

# Output — Markdown
Structure:

## {Sharp, specific title — capture the dominant theme of the week}

[2-3 sentence opening paragraph that sets the tone for the week.]

### Most important
[3-5 bullet points, each starting with the account name in bold. Be specific — names, numbers, deadlines.]

### Accounts that advanced
[Bullet list]

### Accounts at risk
[Bullet list]

### Next week's agenda
[Short paragraph]

### Insight of the week
[The one non-obvious pattern this week reveals. Be opinionated. This is the most important section.]

Keep the entire blog under 450 words. Cut corporate fluff ruthlessly.`,
  },
  {
    key: "weekly-insight",
    name: "Weekly Insight",
    description:
      "Surfaces the single most important pattern of the week (used in the recap callout).",
    systemMessage:
      "You distill a week of sales activity into a single sharp insight. Return one short paragraph in plain text — no markdown, no hedging, no generic platitudes.",
    model: DEFAULT_MODEL,
    deploymentName: DEFAULT_DEPLOYMENT,
    version: "v2.0",
    enabled: true,
    temperature: 0.4,
    maxTokens: 400,
    lastRunDaysAgo: 1,
    tokensLastRun: 320,
    outputFormat: OUTPUT_FORMAT.Text,
    outputSchema: null,
    expectedVariables: [
      {
        name: "week",
        type: "object",
        description: "Week context: number, year.",
      },
      {
        name: "stats",
        type: "object",
        description:
          "Week stats: interactions, risk, pending, bandChanges[] (accountName, from, to).",
      },
      {
        name: "topAccounts",
        type: "array",
        description:
          "Top accounts by touches. Each: name, touches, riskCount.",
      },
    ],
    content: `# Week
Week {{ week.number }} · {{ week.year }}

# Activity summary
- Interactions: {{ stats.interactions }}
- Risk signals: {{ stats.risk }}
- Pending NBAs: {{ stats.pending }}
{% if stats.bandChanges.size > 0 %}- Accounts that moved health band:
{% for change in stats.bandChanges %}  - {{ change.accountName }}: {{ change.from }} → {{ change.to }}
{% endfor %}{% endif %}

# Top accounts this week (by touches)
{% for acc in topAccounts limit:5 %}- {{ acc.name }} ({{ acc.touches }} touches{% if acc.riskCount > 0 %}, {{ acc.riskCount }} risk signal{{ acc.riskCount | pluralize: "", "s" }}{% endif %})
{% endfor %}

# Output — Plain text, 2-3 sentences max
Return a single paragraph (no markdown) that names ONE non-obvious pattern from the week. It should:
- Be specific (cite an account or a number)
- Be opinionated (recommend a stance, not a description)
- Be brief (under 60 words)

Avoid:
- Generic platitudes
- Summarising what already happened (that's the blog's job)
- Hedging language ("might", "perhaps", "consider whether")

Examples of good output:
"Three of the six pending NBAs are tied to competitor pricing pressure on single-aisle. That's a pattern, not a coincidence — pricing strategy needs a hard look at the next leadership meeting."

"Aeroméxico's silence after the counter-proposal matches the LATAM 2024 pattern exactly. Recommend a CEO-to-CEO escalation this week before it becomes a Q3 problem."`,
  },
];

console.log("=== Deleting existing prompts ===");
const existing = await listIds("csp_aiprompts", "csp_aipromptid");
console.log(`  ${existing.length} prompt(s) to delete`);
for (const id of existing) {
  await deleteRecord("csp_aiprompts", id);
}
console.log("  done.\n");

console.log("=== Creating 5 AI prompts ===");
for (const p of PROMPTS) {
  const payload = {
    csp_name: p.name,
    csp_promptkey: p.key,
    csp_description: p.description,
    csp_systemmessage: p.systemMessage,
    csp_content: p.content,
    csp_expectedvariables: JSON.stringify(p.expectedVariables),
    csp_outputformat: p.outputFormat,
    csp_model: p.model,
    csp_deploymentname: p.deploymentName,
    csp_version: p.version,
    csp_enabled: p.enabled,
    csp_lastrunon: daysAgo(p.lastRunDaysAgo),
    csp_tokenslastrun: p.tokensLastRun,
    csp_temperature: p.temperature,
    csp_maxtokens: p.maxTokens,
  };
  if (p.outputSchema) {
    payload.csp_outputschema = JSON.stringify(p.outputSchema, null, 2);
  }
  await createRecord("csp_aiprompts", payload);
  console.log(`  ✓ ${p.name} (${p.key})`);
  await sleep(50);
}

console.log("\nDone. 5 prompts seeded.");
