/**
 * Reseed only the weekly recaps with a coherent 6-week narrative arc.
 *
 * Why this exists: the original seed loops over a generic template and
 * produces 4 near-identical recaps. For the summit demo we want the
 * gallery to feel real — each week telling a distinct story that builds
 * on the prior week.
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

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

// ─── ISO week helpers ────────────────────────────────────────────────────

function getISOWeek(d) {
  const target = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  return Math.ceil(((target - yearStart) / 86400000 + 1) / 7);
}

function isoWeekDates(year, week) {
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const day = jan4.getUTCDay() || 7;
  const monday = new Date(jan4);
  monday.setUTCDate(jan4.getUTCDate() - day + 1 + (week - 1) * 7);
  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);
  return { monday, sunday };
}

// ─── 6-week narrative arc ───────────────────────────────────────────────

const RECAPS = [
  // Oldest first → newest last so they show on the gallery in date desc
  {
    weeksAgo: 5,
    blogTitle: "Season opens quiet — relationship-building over deal-making",
    insight:
      "Three of the four key accounts re-engaged this week showed renewed interest in sustainability-linked terms. That's a thread worth pulling.",
    highlights: [
      "ANA opened the door on A350-1000 pitch for 2027",
      "Lufthansa CFO requested TCO model for upcoming renewal",
      "easyJet confirmed acceptance of A320neo extension",
      "Quiet on the wide-body book — no movement on signings",
    ],
    blogContent: `## Season opens quiet — relationship-building over deal-making

The portfolio enters the new lease cycle from a position of strength. **No urgent fires, no signed deals either** — this week was about laying the groundwork for the back half of the year.

### Most important
- **ANA opened the door** on the A350-1000 pitch for 2027 deliveries. Their board decision is 4–6 weeks out. Reference list from existing operators requested.
- **Lufthansa CFO** asked for a full TCO model — including maintenance reserves — before the renewal discussion proceeds. This is a serious signal: they're benchmarking.
- **easyJet** confirmed the 5-aircraft A320neo extension. Quiet win, but a clean one.

### Accounts that advanced
- ANA — pitch landed, references requested
- Lufthansa — TCO model on critical path
- easyJet — extension confirmed

### Accounts at risk
- None acute this week. Wizz Air's FX exposure was raised again by their CFO but nothing actionable yet.

### Next week's agenda
The TCO model for Lufthansa is the long-pole item. We also need to prepare A220 operator references for the ANA proposal.

### Insight of the week
Three of the four key accounts re-engaged this week showed renewed interest in **sustainability-linked terms**. That's a thread worth pulling at the next sales leadership meeting — possibly we're under-pricing this dimension.`,
    duration: 247,
    tokens: 3120,
  },
  {
    weeksAgo: 4,
    blogTitle: "Pricing pressure surfaces — Ryanair leads, others watching",
    insight:
      "When Ryanair pushes back this hard on price, single-aisle competitors take notice. Expect Wizz and LATAM to ask the same questions within 2-3 weeks.",
    highlights: [
      "Ryanair flagged A320neo lease rate 8% above competitor quote",
      "IAG remarketing package terms finalised — signing imminent",
      "ANA reference list delivered, board meeting set for week 21",
      "Wizz CFO raised FX exposure for the second time this month",
    ],
    blogContent: `## Pricing pressure surfaces — Ryanair leads, others watching

The first **real friction** of the cycle came this week. Ryanair set the tone, but the conversations happening behind closed doors at other LCCs suggest this won't be an isolated push.

### Most important
- **Ryanair** went on the offensive on the A320neo extension. KKR Aviation reportedly quoted 8% below our rate. They want a revised proposal or they'll walk to KKR.
- **IAG** finalised terms on the 5-aircraft remarketing package. Signing expected mid-month. Six weeks of negotiation finally paying off.
- **ANA** reference list delivered. Board meeting now formally on the calendar for week 21.

### Accounts that advanced
- IAG — package terms agreed, signing imminent
- ANA — references in hand, board decision approaching

### Accounts at risk
- **Ryanair** — pricing pressure, competitor active, deadline Friday
- **Wizz Air** — FX exposure raised twice this month; their CFO is clearly building a case

### Next week's agenda
Revised Ryanair proposal is the critical path. Aeroméxico hasn't been heard from in two weeks — worth a proactive check-in.

### Insight of the week
When Ryanair pushes back this hard on price, **single-aisle competitors take notice**. Expect Wizz and LATAM to ask the same questions within 2-3 weeks. We should pre-empt with a refreshed competitive narrative.`,
    duration: 268,
    tokens: 3340,
  },
  {
    weeksAgo: 3,
    blogTitle: "Aeroméxico flagged red — payment delays meet competitor activity",
    insight:
      "Health scoring caught Aeroméxico before we did. The signal mix (late payments + competitor presence + silence) is exactly the pattern that preceded the LATAM near-loss in 2024.",
    highlights: [
      "Aeroméxico moved to Red band — late payments x2 + KKR active in market",
      "IAG package signed (5 aircraft) — €240M ACV",
      "Lufthansa TCO model delivered; CFO response pending",
      "Ryanair revised proposal under review, deadline extended to next Thursday",
    ],
    blogContent: `## Aeroméxico flagged red — payment delays meet competitor activity

This week the **health scoring model earned its keep**. It flagged Aeroméxico into the Red band before any of us would have caught it manually. Two late payments last quarter, plus active KKR Aviation conversations in market, plus radio silence on our last three outreach attempts.

### Most important
- **Aeroméxico** moved into Red. The trigger was the combination of late payments and competitor presence — not any single signal. Worth a proactive call from senior leadership this week.
- **IAG** signed the 5-aircraft package. **€240M ACV**. Six weeks of work, clean execution.
- **Lufthansa** TCO model delivered on Tuesday. CFO response expected by Friday. This is the gating item for the renewal.

### Accounts that advanced
- IAG — signed (5 a/c)
- ANA — board meeting prep on track

### Accounts at risk
- **Aeroméxico** — Red band, all signals flashing
- **Ryanair** — extended their deadline, still on the fence

### Next week's agenda
Aeroméxico recovery is the priority. CFO-to-CFO call this week if possible. Continue Lufthansa decision support.

### Insight of the week
Health scoring caught Aeroméxico **before we did**. The signal mix (late payments + competitor presence + silence) is exactly the pattern that preceded the **LATAM near-loss in 2024**. The model is doing its job — we should be acting on Red flags within 48 hours of them firing.`,
    duration: 295,
    tokens: 3680,
  },
  {
    weeksAgo: 2,
    blogTitle: "Lufthansa renewal on the runway — but Ryanair still circling",
    insight:
      "We're back-to-back on two of the largest renewals of the year. Both are decision-stage. Resourcing-wise, the team can't afford to lose focus on either.",
    highlights: [
      "Lufthansa CFO approved MR rate adjustment — paves way for renewal",
      "Aeroméxico CFO-to-CFO call held — tense but constructive",
      "Ryanair still undecided; meeting scheduled for next Wednesday",
      "LATAM made first contact request in 6 months — likely fleet refresh",
    ],
    blogContent: `## Lufthansa renewal on the runway — but Ryanair still circling

Two of our largest renewals are now in the **decision corridor**. This is the moment when account managers earn their pay.

### Most important
- **Lufthansa CFO** approved the maintenance reserve rate adjustment. This was the gating item. Renewal contract drafting begins next week. **€520M over 8 years.**
- **Aeroméxico** — the CFO-to-CFO call happened. Tense, but constructive. They are not gone. They want a restructured payment schedule and 18 months of relief. Counter-proposal due Monday.
- **Ryanair** still undecided. Meeting confirmed for Wednesday. KKR has not closed them either — there's still a window.

### Accounts that advanced
- Lufthansa — renewal drafting starts
- Aeroméxico — back in the conversation
- LATAM — first contact in 6 months, fleet refresh likely

### Accounts at risk
- **Ryanair** — still on the edge
- **Wizz Air** — silent again this week, third week of light contact

### Next week's agenda
Lufthansa renewal contract first draft. Aeroméxico restructure counter-proposal. Ryanair meeting Wednesday. LATAM discovery call.

### Insight of the week
We're back-to-back on **two of the largest renewals of the year**. Both are decision-stage. Resourcing-wise, the team can't afford to lose focus on either — and we're about to add LATAM into the active book. Worth flagging to leadership for backfill conversation.`,
    duration: 312,
    tokens: 3920,
  },
  {
    weeksAgo: 1,
    blogTitle: "Ryanair lost — but IAG and Lufthansa more than make up for it",
    insight:
      "Three of the six pending NBAs are now tied to competitor pricing pressure on single-aisle. That's a pattern, not a coincidence — pricing strategy needs a hard look at the next leadership meeting.",
    highlights: [
      "Ryanair confirmed walking to KKR — 3 a/c lost, ~€90M ACV",
      "Lufthansa renewal contract first draft delivered",
      "ANA board decision moved up to next week",
      "LATAM fleet refresh tender announced — we're invited",
    ],
    blogContent: `## Ryanair lost — but IAG and Lufthansa more than make up for it

We lost Ryanair. **3 aircraft, ~€90M ACV gone to KKR Aviation.** It hurts. But this week also brought a clean draft of the Lufthansa renewal and confirmation that ANA is moving faster than expected.

### Most important
- **Ryanair** walked. Final price gap was 6% — they took the KKR offer. We knew this was a possibility two weeks ago when we declined to chase below 92% margin. Discipline holding, but the loss is real.
- **Lufthansa renewal** — first contract draft delivered. €520M, 8 years, 14 aircraft. Legal review next week.
- **ANA** — board meeting moved up to next week. Likely positive given the reference calls landed well.

### Accounts that advanced
- Lufthansa — contract drafting
- ANA — decision imminent
- LATAM — tender invitation received

### Accounts at risk
- **Aeroméxico** still in counter-proposal cycle; CFO responsive but slow
- **Wizz Air** silence is now four weeks; needs an outreach push

### Next week's agenda
Lufthansa legal review. ANA board outcome. Aeroméxico restructure response. LATAM tender prep.

### Insight of the week
Three of the six pending NBAs are now tied to **competitor pricing pressure on single-aisle**. That's a pattern, not a coincidence. Ryanair lost was the symptom — the underlying issue is that we are systematically being priced 6-8% above KKR on single-aisle. Pricing strategy needs a hard look at the next leadership meeting.`,
    duration: 334,
    tokens: 4180,
  },
  {
    weeksAgo: 0,
    blogTitle: "Two wins land, one risk solidifies — portfolio holds steady",
    insight:
      "This was the most productive week of the cycle so far — but Aeroméxico's quiet response to the counter-proposal is the silent risk. The accounts that go cold are the ones we lose.",
    highlights: [
      "ANA board approved A350-1000 — 2 aircraft, 2027 delivery",
      "Lufthansa renewal contract signed and countersigned",
      "Aeroméxico went quiet again after counter-proposal sent",
      "LATAM tender preparation in progress, deadline 3 weeks",
    ],
    blogContent: `## Two wins land, one risk solidifies — portfolio holds steady

**The two biggest deals of the cycle closed this week.** Lufthansa renewed. ANA committed. Combined ACV well over **€700M**. The portfolio is in a fundamentally stronger position than 6 weeks ago.

### Most important
- **ANA** board approved the A350-1000 deal. 2 aircraft, 2027 delivery, optionality on a third. Cleanest pitch-to-close in 18 months.
- **Lufthansa** renewal contract signed and countersigned. €520M over 8 years. 14 aircraft locked in.
- **Aeroméxico** received our counter-proposal Monday. **No response since Wednesday.** This is the silent risk now.

### Accounts that advanced
- ANA — signed
- Lufthansa — renewed
- LATAM — tender prep underway

### Accounts at risk
- **Aeroméxico** — silent. The pattern matches LATAM 2024 more closely each week.
- **Wizz Air** — five weeks of light contact, need a senior touch

### Next week's agenda
Aeroméxico proactive outreach (CFO-level). LATAM tender response. Wizz Air recovery plan.

### Insight of the week
This was the most productive week of the cycle so far — but **Aeroméxico's quiet response to the counter-proposal is the silent risk**. The accounts that go cold are the ones we lose. Wizz Air is showing the same pattern with three weeks earlier on the curve. Recommend a Red-band review specifically on response cadence as a leading indicator, not just outcome metrics.`,
    duration: 358,
    tokens: 4520,
  },
];

// ─── execution ───────────────────────────────────────────────────────────

console.log("=== Deleting existing recaps ===");
const existing = await listIds("csp_weeklyrecaps", "csp_weeklyrecapid");
console.log(`  ${existing.length} recap(s) to delete`);
for (const id of existing) {
  await deleteRecord("csp_weeklyrecaps", id);
}
console.log("  done.\n");

console.log("=== Creating 6 narrative recaps ===");
const now = new Date();
for (const r of RECAPS) {
  const weekDate = new Date(now);
  weekDate.setDate(weekDate.getDate() - r.weeksAgo * 7);
  const weekNum = getISOWeek(weekDate);
  const year = weekDate.getFullYear();
  const { monday } = isoWeekDates(year, weekNum);

  const payload = {
    csp_name: `Week ${weekNum} - ${year}`,
    csp_weeknumber: weekNum,
    csp_year: year,
    csp_blogtitle: r.blogTitle,
    csp_blogcontent: r.blogContent,
    csp_audiourl: `https://example.blob.core.windows.net/podcasts/week-${year}-${weekNum}.mp3`,
    csp_audioduration: r.duration,
    csp_coverurl: `https://example.blob.core.windows.net/covers/week-${year}-${weekNum}.png`,
    csp_highlights: JSON.stringify(r.highlights),
    csp_insight: r.insight,
    csp_tokensused: r.tokens,
    overriddencreatedon: monday.toISOString(),
  };

  await createRecord("csp_weeklyrecaps", payload);
  console.log(`  ✓ Week ${weekNum} ${year} — "${r.blogTitle}"`);
  await sleep(50);
}

console.log("\nDone. 6 recaps reseeded.");
