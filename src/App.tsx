import { useEffect, useMemo, useState, type ReactNode } from "react";
import { Apps20Regular as Blocks } from "@fluentui/react-icons/svg/apps";
import { ArrowRight20Regular as ArrowRight } from "@fluentui/react-icons/svg/arrow-right";
import { ArrowSync20Regular as RefreshCw } from "@fluentui/react-icons/svg/arrow-sync";
import { Bot20Regular as Bot } from "@fluentui/react-icons/svg/bot";
import { Branch20Regular as GitBranch } from "@fluentui/react-icons/svg/branch";
import { Braces20Filled, Braces20Regular as Braces } from "@fluentui/react-icons/svg/braces";
import { Checkmark20Regular as Check } from "@fluentui/react-icons/svg/checkmark";
import { CheckmarkCircle20Regular as CheckCircle } from "@fluentui/react-icons/svg/checkmark-circle";
import { ChevronRight20Regular as ChevronRight } from "@fluentui/react-icons/svg/chevron-right";
import { Clock20Regular as Clock } from "@fluentui/react-icons/svg/clock";
import { Database20Regular as Database } from "@fluentui/react-icons/svg/database";
import { Flow20Regular as Flow } from "@fluentui/react-icons/svg/flow";
import { Grid20Filled, Grid20Regular as Grid } from "@fluentui/react-icons/svg/grid";
import { History20Regular as History } from "@fluentui/react-icons/svg/history";
import { Link20Regular as LinkIcon } from "@fluentui/react-icons/svg/link";
import { Play20Regular as Play } from "@fluentui/react-icons/svg/play";
import { Search20Regular as Search } from "@fluentui/react-icons/svg/search";
import { Settings20Regular as Settings } from "@fluentui/react-icons/svg/settings";
import { Sparkle20Regular as Sparkles } from "@fluentui/react-icons/svg/sparkle";
import { Target20Regular as Target } from "@fluentui/react-icons/svg/target";
import { WindowConsole20Regular as Code } from "@fluentui/react-icons/svg/window-console";
import { Link, NavLink, Route, Routes, useNavigate, useParams } from "react-router-dom";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { VisualFilterBuilder } from "@/components/VisualFilterBuilder";
import { cn } from "@/lib/utils";
import { buildConfigurationPayload, buildContextFetchXml, buildDraftConfigurationPayload, buildPublishSignal, buildRecordSelectionFetchXml, compileRecipe, getTargetEntityIdentity, validateSummaryConfiguration, type SummaryConfigurationDraft } from "@/lib/summaryConfiguration";
import { useCreateStudioConfiguration, useStudioRuntime, useUpdateStudioConfiguration, useUpdateStudioPrompt, type StudioAccount, type StudioConfiguration, type StudioPrompt } from "@/hooks/useStudioRuntime";
import { useSystemViews } from "@/hooks/useSystemViews";
import { useUserViews } from "@/hooks/useUserViews";

const LOGO = "https://99e46348.delivery.rocketcdn.me/wp-content/uploads/2024/10/creativity-spark-blanco.svg";
const MARK = "/creativity-spark-mark.png";

const navItems = [
  { to: "/", label: "Configurations", icon: Grid, activeIcon: Grid20Filled, end: true },
  { to: "/configurations/account-operations", label: "Design summary", icon: Braces, activeIcon: Braces20Filled },
  { to: "/runs", label: "Runs", icon: History, activeIcon: History },
];

function Shell({ children }: { children: ReactNode }) {
  const { data } = useStudioRuntime();
  const live = data?.live ?? false;
  return (
    <div className="studio-shell">
      <a href="#main-content" className="skip-link">Skip to main content</a>
      <aside className="studio-sidebar">
        <div className="px-5 pt-6">
          <img src={LOGO} alt="Creativity Spark" width="188" height="74" className="h-auto w-[170px]" />
          <div className="mt-5 flex items-center gap-2 border-t border-white/10 pt-4">
            <img src={MARK} alt="" width="26" height="26" className="size-6 object-contain" />
            <div><p className="text-[13px] font-semibold text-white">Summary Studio</p><p className="text-[9px] font-semibold uppercase tracking-[.18em] text-white/38">Power Platform compiler</p></div>
          </div>
        </div>
        <nav className="mt-8 flex-1 border-t border-white/10 px-3 pt-7" aria-label="Primary navigation">
          <p className="micro-label mb-2 px-3 text-white/28">Workspace</p>
          {navItems.map(({ to, label, icon: Icon, activeIcon: ActiveIcon, end }) => (
            <NavLink key={to} to={to} end={end} className={({ isActive }) => cn("studio-nav", isActive && "studio-nav-active")}>
              {({ isActive }) => <><span className="nav-icon">{isActive ? <ActiveIcon /> : <Icon />}</span><span className="flex-1">{label}</span>{isActive && <i />}</>}
            </NavLink>
          ))}
        </nav>
        <div className="mx-4 mb-4 rounded-xl border border-white/10 bg-white/[.035] p-3">
          <div className="flex items-center gap-2 text-[10px] text-white/50"><span className="size-1.5 rounded-full bg-emerald-400" />Demo environment</div>
          <p className="mt-2 font-mono text-[9px] text-white/34">spark-tools-dev.crm4</p>
        </div>
        <div className="m-4 mt-0 border-t border-white/10 pt-4">
          <div className="flex items-center gap-3"><Avatar className="size-8 border border-white/10"><AvatarFallback className="bg-brand-blue text-[10px] font-semibold text-white">OF</AvatarFallback></Avatar><div className="min-w-0 flex-1"><p className="truncate text-[11px] font-medium text-white">Oscar Fuentes</p><p className="text-[9px] text-white/35">Maker · administrator</p></div><Settings className="size-4 text-white/30" /></div>
        </div>
      </aside>
      <main id="main-content" className="min-w-0 flex-1 overflow-x-hidden">
        <header className="studio-topbar">
          <div className="flex items-center gap-3 text-xs text-muted-foreground"><span className="font-medium text-foreground">AI Summaries in Dataverse</span><ChevronRight className="size-3.5" /><span>Above and Beyond</span></div>
          <div className="flex items-center gap-3"><span className="hidden items-center gap-2 text-[10px] text-muted-foreground md:flex"><span className="size-1.5 rounded-full bg-emerald-500" />{live ? "Dataverse connected" : "Local fallback"}</span><Button variant="outline" size="sm"><Search data-icon="inline-start" />Search</Button></div>
        </header>
        {children}
      </main>
    </div>
  );
}

function PageHeader({ eyebrow, title, description, actions }: { eyebrow: string; title: string; description: string; actions?: ReactNode }) {
  return <div className="page-header"><div><p className="micro-label text-brand-blue">{eyebrow}</p><h1>{title}</h1><p>{description}</p></div>{actions && <div className="flex shrink-0 items-center gap-2">{actions}</div>}</div>;
}

function Overview() {
  const { data, refetch, isFetching } = useStudioRuntime();
  const runtime = data;
  const configurations = runtime?.configurations ?? [];
  const published = configurations.filter((item) => item.status === 100000001).length;
  return <Shell><div className="studio-page">
    <section className="compiler-hero compiler-hero-compact">
      <div className="hero-compact-copy"><p className="micro-label text-brand-cyan">Summary Studio · declarative configuration</p><h1>Configure AI summaries</h1><p>Define context, prompt, destination, and execution in a single Dataverse recipe.</p></div>
      <div className="hero-compact-actions"><Button asChild className="bg-brand-cyan text-brand-navy hover:bg-brand-cyan/90"><Link to="/configurations/new"><Sparkles data-icon="inline-start" />New configuration</Link></Button>{configurations[0] && <Link to={`/configurations/${configurations[0].id}`} className="hero-link">Open recipe <ArrowRight /></Link>}</div>
      <div className="compiler-pipeline" aria-label="Summary recipe and Power Platform components"><span><Braces /><b>Recipe</b><small>context + rules</small></span><ChevronRight /><span><Database /><b>Dataverse</b><small>configuration</small></span><ChevronRight /><span><Bot /><b>AI Prompt</b><small>instructions</small></span><ChevronRight /><span><Flow /><b>Backend</b><small>materialization</small></span></div>
    </section>

    <section className="stat-ribbon" aria-label="Studio summary"><div><b>{published}</b><span>published configurations</span></div><div><b>{runtime?.cacheCount ?? 0}</b><span>available summaries</span></div><div><b>{runtime?.runs.length ?? 0}</b><span>recorded runs</span></div><div><b>{runtime?.accounts.length ?? 0}</b><span>demo records</span></div></section>

    <section className="mt-6">
      <div className="section-heading"><div><p className="micro-label text-muted-foreground">Catalog</p><h2>Summary configurations</h2></div><div className="flex gap-2"><Button variant="outline" size="sm" onClick={() => refetch()} disabled={isFetching}><RefreshCw data-icon="inline-start" />{isFetching ? "Refreshing" : "Refresh"}</Button><Button asChild size="sm"><Link to="/configurations/new"><Sparkles data-icon="inline-start" />Create summary</Link></Button></div></div>
      <div className="config-table">
        <div className="config-table-head"><span>Configuration</span><span>Source and destination</span><span>Trigger</span><span>Activity</span><span>Status</span><span /></div>
        {configurations.map((item, index) => <Link className="config-row" to={`/configurations/${item.id}`} key={item.id || item.name}>
          <div className="flex min-w-0 items-center gap-3"><span className={cn("entity-icon", index === 1 && "violet", index === 2 && "green", index === 3 && "amber")}><Database /></span><span className="min-w-0"><b>{item.name}</b><small>{item.promptName || "AI Prompt"}</small></span></div>
          <div><b>{item.entity === "account" ? "Account" : item.entity}</b><small>{item.outputEntity}.{item.outputField}</small></div><div><b>{item.mode === 100000000 ? "On update" : "Scheduled"}</b><small>{item.flowName || "Flow pending"}</small></div><div><b>{runtime?.runs.length ?? 0} runs</b><small>{item.lastRun ? "Activity recorded" : "Not run"}</small></div><div><Badge variant="outline" className={item.status === 100000001 ? "status-live" : "status-draft"}>{item.status === 100000001 ? "Published" : "Draft"}</Badge></div><ChevronRight className="size-4 text-muted-foreground" />
        </Link>)}
      </div>
    </section>
    <section className="architecture-note"><span><Blocks /></span><div><b>One engine, multiple patterns</b><p>Configurations are stored in Dataverse and materialized through governed Power Automate templates: event-driven, scheduled, or on demand.</p></div><code>csp_aisummaryconfig</code><div className="extension-marker"><Sparkles /><span><small>Extension ready</small><b>Foundry evaluations</b></span></div></section>
  </div></Shell>;
}

const steps = [
  { title: "Data and context", stage: "Context", subtitle: "What it should read", icon: Database },
  { title: "Prompt and model", stage: "Prompt", subtitle: "How it should reason", icon: Bot },
  { title: "Destination", stage: "Destination", subtitle: "Where it should write", icon: Target },
  { title: "Execution", stage: "Automation", subtitle: "When it should run", icon: Flow },
  { title: "Review and publish", stage: "Publish", subtitle: "What will be generated", icon: CheckCircle },
];

function LabeledField({ label, children, hint }: { label: string; children: ReactNode; hint?: string }) {
  return <label className="studio-field"><span>{label}</span>{children}{hint && <small>{hint}</small>}</label>;
}

type PaneKind = "table" | "execution" | "prompt" | "model" | "fields" | "records" | "relationships" | "contextQuery" | "outputTable" | "outputField";

type BuilderDraft = SummaryConfigurationDraft;

function draftFromConfiguration(config?: StudioConfiguration, prompts: StudioPrompt[] = []): BuilderDraft {
  const prompt = prompts.find((item) => item.id === config?.promptId) ?? prompts[0];
  const entity = config?.entity ?? "account";
  const sourceFields = config?.sourceFields.length ? config.sourceFields : ["name"];
  const relationships = config?.relationships ?? [];
  const maxRecords = config?.maxRecords ?? 5000;
  return {
    name: config?.name ?? "New summary", entity,
    sourceFields,
    outputEntity: config?.outputEntity ?? "account", outputField: config?.outputField ?? "csp_aisummary",
    model: config?.model ?? prompt?.model ?? "gpt-4.1-mini", mode: config?.mode ?? 100000000,
    promptId: prompt?.id ?? "", promptName: prompt?.name ?? "Select AI Prompt",
    promptKey: prompt?.key ?? "", promptContent: prompt?.content ?? "", flowName: config?.flowName ?? "",
    relationships, inputMappings: config?.inputMappings ?? {},
    triggerColumns: config?.triggerColumns.length ? config.triggerColumns : ["name", "revenue", "description", "primarycontactid"],
    triggerType: "dataverse.update", preserveHistory: false, saveMetadata: true,
    queryMode: config?.queryMode ?? 100000002, maxRecords,
    fetchXml: config?.fetchXml || buildRecordSelectionFetchXml(entity, maxRecords),
    relatedFetchXml: config?.relatedFetchXml || buildContextFetchXml(entity, sourceFields, relationships),
  };
}

function DataStep({ onEdit, draft, accounts }: { onEdit: (pane: PaneKind) => void; draft: BuilderDraft; accounts: StudioAccount[] }) {
  const entityLabel = draft.entity === "account" ? "Accounts" : draft.entity === "incident" ? "Cases" : "Opportunities";
  const relationshipCount = draft.relationships.length;
  const recordSelectionLabel = draft.queryMode === 100000000 ? "System view" : draft.queryMode === 100000001 ? "Personal view" : "Filters or FetchXML";
  const relatedMode = Number(draft.relationships[0]?.selectionMode ?? 100000002);
  const relatedSelectionLabel = relatedMode === 100000000 ? "System view" : relatedMode === 100000001 ? "Personal view" : "Filters or FetchXML";
  return <div className="editor-section data-context-step"><SectionIntro number="01" title="Data and context" text="Configure the record set first, then define the Dataverse context sent to the prompt for each record." />
    <div className="configuration-sections">
      <section className="configuration-section" aria-label="1. Records to summarize">
        <header><span>1</span><div><h3>Records to summarize</h3><p>Define which Dataverse rows enter the generated flow.</p></div></header>
        <div className="configuration-list">
          <article><span className="configuration-icon"><Database /></span><div className="configuration-label"><b>Source table</b><code>csp_targetentity</code></div><div className="configuration-value"><b>{entityLabel}</b><code>{draft.entity}</code></div><Button type="button" variant="outline" size="sm" aria-label="Change source table" onClick={() => onEdit("table")}>Change</Button></article>
          <article><span className="configuration-icon"><Search /></span><div className="configuration-label"><b>Record selection</b><code>csp_fetchxml</code></div><div className="configuration-value"><b>{recordSelectionLabel}</b><small>{accounts.length} demo records</small><small>Maximum {draft.maxRecords.toLocaleString("en-US")}</small></div><Button type="button" variant="outline" size="sm" aria-label="Configure record FetchXML" onClick={() => onEdit("records")}>Configure</Button></article>
          <article><span className="configuration-icon"><Blocks /></span><div className="configuration-label"><b>Summary mode</b><code>csp_mode</code></div><div className="configuration-value"><b>{draft.mode === 100000000 ? "One summary per record" : "Consolidated summary"}</b><code>{draft.mode === 100000000 ? "PerRecord" : "Aggregate"}</code></div><Button type="button" variant="outline" size="sm" aria-label="Change summary mode" onClick={() => onEdit("execution")}>Change</Button></article>
        </div>
      </section>

      <section className="configuration-section" aria-label="2. Context for each record">
        <header><span>2</span><div><h3>Context for each record</h3><p>Control exactly what is passed to the selected AI Prompt.</p></div></header>
        <div className="configuration-list">
          <article><span className="configuration-icon"><Database /></span><div className="configuration-label"><b>Source fields</b><code>csp_sourcefields</code></div><div className="configuration-value configuration-tokens">{draft.sourceFields.map((field) => <code key={field}>{field}</code>)}</div><Button type="button" variant="outline" size="sm" aria-label="Configure source fields" onClick={() => onEdit("fields")}>Select fields</Button></article>
          <article><span className="configuration-icon"><GitBranch /></span><div className="configuration-label"><b>Related records</b><code>csp_relationships</code></div><div className="configuration-value"><b>{relationshipCount ? relatedSelectionLabel : "No related records"}</b><code>{relationshipCount ? `${String(draft.relationships[0]?.entity ?? "activitypointer")} · 1:N` : "none"}</code></div><Button type="button" variant="outline" size="sm" aria-label="Configure related records" onClick={() => onEdit("relationships")}>Configure</Button></article>
          <article><span className="configuration-icon"><Code /></span><div className="configuration-label"><b>Context query</b><code>csp_relatedfetchxml</code></div><div className="configuration-value"><b>FetchXML per record</b><small>{draft.sourceFields.length} fields · {relationshipCount} related table</small></div><Button type="button" variant="outline" size="sm" aria-label="Configure context FetchXML" onClick={() => onEdit("contextQuery")}>Edit FetchXML</Button></article>
        </div>
      </section>
    </div>
    <div className="configuration-result"><CheckCircle /><span><b>Compiled context</b><small>For every matching <code>{draft.entity}</code>, the flow runs <code>csp_relatedfetchxml</code> and passes its result to the prompt.</small></span><Button type="button" variant="outline" size="sm" onClick={() => onEdit("contextQuery")}>Review query</Button></div>
  </div>;
}

function PromptStep({ onEdit, draft, onSavePrompt, savingPrompt }: { onEdit: (pane: PaneKind) => void; draft: BuilderDraft; onSavePrompt: (content: string) => void; savingPrompt: boolean }) {
  const [assistantOpen, setAssistantOpen] = useState(false);
  const [refined, setRefined] = useState(false);
  const [promptContent, setPromptContent] = useState(draft.promptContent);

  return <div className="editor-section"><SectionIntro number="02" title="Prompt and model" text="Keep the maker experience simple without hiding the architecture and governance decisions that matter." />
    <div className="grid gap-4 md:grid-cols-[1fr_220px]"><LabeledField label="AI Prompt"><button type="button" className="field-control w-full text-left" aria-label="Edit AI Prompt" onClick={() => onEdit("prompt")}><Bot /><span><b>{draft.promptName}</b><code>{draft.promptKey}</code></span><ChevronRight /></button></LabeledField><LabeledField label="Model"><button type="button" className="field-control compact w-full text-left" aria-label="Edit model" onClick={() => onEdit("model")}><Sparkles /><span><b>{draft.model.replace("gpt-", "GPT-").replace("-mini", " mini")}</b><code>EU Data Zone</code></span></button></LabeledField></div>
    <LabeledField label="Prompt instructions" hint="Changes are saved to the selected AI Prompt record in Dataverse."><div className="rounded-xl border border-border bg-white p-3"><textarea className="min-h-28 w-full resize-none bg-transparent font-mono text-sm leading-6 outline-none" value={promptContent} onChange={(event) => setPromptContent(event.target.value)} aria-label="Prompt instructions" /><div className="mt-2 flex items-center justify-between border-t border-border pt-2"><code className="rounded-md bg-cyan-50 px-2 py-1 text-xs text-brand-blue">{"{{account_context}}"}</code><Button type="button" size="sm" variant="outline" disabled={savingPrompt || !draft.promptId || promptContent === draft.promptContent} onClick={() => onSavePrompt(promptContent)}>{savingPrompt ? "Saving prompt…" : "Save prompt"}</Button></div></div></LabeledField>
    <div className="prompt-assistant-trigger"><div><span><Sparkles /></span><p><b>Design a more consistent prompt</b><small>Analyze the selected objective, audience, and Dataverse context.</small></p></div><Button type="button" variant="outline" size="sm" onClick={() => setAssistantOpen(!assistantOpen)}><Sparkles data-icon="inline-start" />Help me improve this prompt</Button></div>
    {assistantOpen && <section className="prompt-assistant" aria-label="Prompt assistant"><header><div><span><Sparkles /></span><p className="micro-label">Prompt assistant</p><h3>Improvements based on your recipe</h3></div><button type="button" aria-label="Close prompt assistant" onClick={() => setAssistantOpen(false)}>×</button></header><div className="assistant-context"><Database /><span><b>Context detected · Account + activities</b><small>6 fields, 1 relationship, 24 sample records, and an executive audience.</small></span><em>Ready to refine</em></div><div className="assistant-suggestions"><article><span>01</span><div><b>Reduce non-essential context</b><p>Exclude <code>description</code> when no changes exist. Preserve signal while reducing input size.</p></div><strong>−18% estimated tokens</strong></article><article><span>02</span><div><b>Make the output verifiable</b><p>Define fixed sections and explicitly identify unknown data.</p></div><strong>More consistent</strong></article><article><span>03</span><div><b>Adapt tone to the audience</b><p>Prioritize commercial impact, risk, and next action for executive readers.</p></div><strong>More useful</strong></article></div><footer><p><CheckCircle /> The proposal preserves the <code>{"{{account_context}}"}</code> input and the rules against invented information.</p><Button type="button" onClick={() => { setRefined(true); setPromptContent((current) => `${current.trim()}\n\nUse fixed sections. State unknown when information is absent. Prioritize business impact, risk, and the next action.`); }} disabled={refined}><Check data-icon="inline-start" />{refined ? "Refined version applied" : "Apply refined version"}</Button></footer></section>}
    <div className="input-grid"><div><p className="field-caption">Runtime inputs</p><span><Database />{"{{account_context}}"}<small>JSON · compiled from FetchXML</small></span><span><Clock />{"{{generated_at}}"}<small>DateTime · Europe/Madrid</small></span></div><aside><p>Estimate · {refined ? "1,017" : "1,240"} tokens per run</p><b>≈ €{refined ? "0.0007" : "0.0009"}</b><small>{refined ? "Refined version · estimated 18% saving." : "Calculated from the current 24-record sample."}</small></aside></div>
  </div>;
}

function DestinationStep({ draft, onEdit, onChange }: { draft: BuilderDraft; onEdit: (pane: PaneKind) => void; onChange: (changes: Partial<BuilderDraft>) => void }) {
  return <div className="editor-section"><SectionIntro number="03" title="Destination" text="Write the summary back to Dataverse so it can be reused in views, forms, agents, and other automations." />
    <div className="destination-map"><div><Database /><span><small>Source record</small><b>{draft.entity === "account" ? "Account" : draft.entity}</b><code>{draft.entity}.{draft.entity}id</code></span></div><ArrowRight /><div className="active"><Sparkles /><span><small>Generated result</small><b>Executive summary</b><code>{draft.outputEntity}.{draft.outputField}</code></span></div></div>
    <div className="grid gap-4 md:grid-cols-2"><LabeledField label="Destination table"><button type="button" className="field-control w-full text-left" aria-label="Edit destination table" onClick={() => onEdit("outputTable")}><Database /><span><b>{draft.outputEntity === "account" ? "Account" : "Summary cache"}</b><code>{draft.outputEntity}</code></span><ChevronRight /></button></LabeledField><LabeledField label="Destination column"><button type="button" className="field-control w-full text-left" aria-label="Edit destination column" onClick={() => onEdit("outputField")}><Target /><span><b>{draft.outputField === "csp_aisummary" ? "AI summary" : draft.outputField === "description" ? "Description" : "Name"}</b><code>{draft.outputField}</code></span><ChevronRight /></button></LabeledField></div>
    <div className="option-list"><button type="button" aria-pressed={!draft.preserveHistory} className={!draft.preserveHistory ? "selected" : ""} onClick={() => onChange({ preserveHistory: false })}><span className="radio-dot" /><div><b>Overwrite the current summary</b><small>Keep one current result on the source record.</small></div>{!draft.preserveHistory && <Check />}</button><button type="button" aria-pressed={draft.preserveHistory} className={draft.preserveHistory ? "selected" : ""} onClick={() => onChange({ preserveHistory: true })}><span className="radio-dot" /><div><b>Preserve history</b><small>Create a cache entry for every run.</small></div>{draft.preserveHistory && <Check />}</button></div>
    <div className="inline-setting"><div><b>Store generation metadata</b><small>Model, tokens, duration, and configuration version.</small></div><Switch checked={draft.saveMetadata} onCheckedChange={(checked) => onChange({ saveMetadata: checked })} aria-label="Store generation metadata" /></div>
  </div>;
}

function triggerPresentation(type: string, columns: string[]) {
  if (type === "schedule.daily") return { label: "Daily schedule", detail: "Every day · Europe/Madrid", pattern: "Scheduled" };
  if (type === "manual") return { label: "On-demand execution", detail: "Explicit invocation", pattern: "On demand" };
  return { label: "When an account changes", detail: columns.slice(0, 3).join(", "), pattern: "Dataverse event" };
}

function FlowDiagram({ draft }: { draft: BuilderDraft }) {
  const trigger = triggerPresentation(draft.triggerType, draft.triggerColumns);
  return <div className="generated-flow"><div className="flow-node trigger"><span><Clock /></span><div><small>Trigger</small><b>{trigger.label}</b><code>{trigger.detail}</code></div></div><i /><div className="flow-node"><span><Database /></span><div><small>Dataverse</small><b>Query Dataverse</b><code>FetchXML + relationships</code></div></div><i /><div className="flow-node"><span><Bot /></span><div><small>AI Builder</small><b>Run AI Prompt</b><code>{draft.promptKey}</code></div></div><i /><div className="flow-node"><span><Target /></span><div><small>Dataverse</small><b>Save summary</b><code>{draft.outputEntity}.{draft.outputField}</code></div></div></div>;
}

function ExecutionStep({ draft, onChange }: { draft: BuilderDraft; onChange: (changes: Partial<BuilderDraft>) => void }) {
  const trigger = triggerPresentation(draft.triggerType, draft.triggerColumns);
  return <div className="editor-section"><SectionIntro number="04" title="Execution" text="Select a governed pattern. Summary Studio generates the trigger and actions without turning the app into a universal flow designer." />
    <div className="trigger-options"><button type="button" aria-label="When an account changes" aria-pressed={draft.triggerType === "dataverse.update"} className={draft.triggerType === "dataverse.update" ? "selected" : ""} onClick={() => onChange({ triggerType: "dataverse.update" })}><span><GitBranch /></span><div><b>When an account changes</b><small>Only when a field used by the context changes.</small></div>{draft.triggerType === "dataverse.update" && <Check />}</button><button type="button" aria-label="Scheduled" aria-pressed={draft.triggerType === "schedule.daily"} className={draft.triggerType === "schedule.daily" ? "selected" : ""} onClick={() => onChange({ triggerType: "schedule.daily" })}><span><Clock /></span><div><b>Scheduled</b><small>Daily, weekly, or a custom recurrence.</small></div>{draft.triggerType === "schedule.daily" && <Check />}</button><button type="button" aria-label="On demand" aria-pressed={draft.triggerType === "manual"} className={draft.triggerType === "manual" ? "selected" : ""} onClick={() => onChange({ triggerType: "manual" })}><span><Play /></span><div><b>On demand</b><small>From a command or as a child flow.</small></div>{draft.triggerType === "manual" && <Check />}</button></div>
    <div className="change-filter"><div><p className="field-caption">Trigger columns</p><div>{draft.triggerColumns.map(x=><span key={x}>{x}</span>)}</div></div><aside><RefreshCw /><span><b>Loop prevention active</b><small>The output column is excluded from the trigger.</small></span></aside></div>
    <div className="flow-preview"><div className="flex items-center justify-between"><div><p className="field-caption">Compiled definition</p><h3>Flow contract for the backend</h3></div><Badge variant="outline">Pattern · {trigger.pattern}</Badge></div><FlowDiagram draft={draft} /></div>
  </div>;
}

function ReviewStep({ draft, validationErrors }: { draft: BuilderDraft; validationErrors: string[] }) {
  const trigger = triggerPresentation(draft.triggerType, draft.triggerColumns);
  return <div className="editor-section"><SectionIntro number="05" title="Review and publish" text="The recipe is complete. Publishing creates or updates only the assets required by the solution." />
    <div className="review-grid"><ReviewCard icon={<Database />} title="Data" value="Account + activities" detail={`${draft.sourceFields.length} fields · 1 relationship`} /><ReviewCard icon={<Bot />} title="Generation" value={draft.promptName} detail={`${draft.model} · 2 inputs · English`} /><ReviewCard icon={<Target />} title="Destination" value={`${draft.outputEntity}.${draft.outputField}`} detail="Overwrite · metadata enabled" /><ReviewCard icon={<Flow />} title="Execution" value={trigger.label} detail={draft.triggerType === "dataverse.update" ? "Loop prevention prepared" : trigger.detail} /></div>
    <div className="publish-contract"><div><p className="field-caption">Prepared assets</p><div className="asset-list"><span><CheckCircle />1 configuration record <code>csp_aisummaryconfig</code></span><span><CheckCircle />1 versioned definition for the flow generator <code>{draft.flowName || "Generated after publishing"}</code></span><span><LinkIcon />1 binding to an existing AI Prompt <code>{draft.promptKey || "Not selected"}</code></span></div></div><aside aria-live="polite">{validationErrors.length === 0 ? <><p><CheckCircle /> Validation complete</p><small>Valid schema; flow materialization is handled by the backend.</small></> : <><p>Resolve {validationErrors.length} item{validationErrors.length === 1 ? "" : "s"}</p><ul className="mt-2 list-disc space-y-1 pl-4 text-xs">{validationErrors.map((error) => <li key={error}>{error}</li>)}</ul></>}</aside></div>
  </div>;
}

function ReviewCard({ icon, title, value, detail }: { icon: ReactNode; title: string; value: string; detail: string }) { return <div className="review-card"><span>{icon}</span><p>{title}</p><b>{value}</b><small>{detail}</small></div>; }
function SectionIntro({ number, title, text }: { number: string; title: string; text: string }) { return <div className="section-intro"><span>{number}</span><div><h2>{title}</h2><p>{text}</p></div></div>; }

function CompiledRecipe({ step, draft, title = "Recipe v3" }: { step: number; draft: BuilderDraft; title?: string }) {
  const status = ["context ready", "prompt bound", "destination valid", "flow prepared", "ready to publish"][step];
  return <aside className="compiled-panel"><div className="compiled-head"><div><p className="micro-label text-brand-cyan">Compiled view</p><h2>{title}</h2></div><span><i />{status}</span></div>
    <div className="recipe-line"><span>OBJECTIVE</span><b>{draft.name}</b></div>
    <div className="recipe-code"><div><span>01</span><code>source</code><b>table · {draft.entity}</b></div><div><span>02</span><code>fields</code><b>{draft.sourceFields.length} selected</b></div><div><span>03</span><code>relation</code><b>activitypointer</b></div><div><span>04</span><code>prompt</code><b>{draft.promptKey}</b></div><div><span>05</span><code>model</code><b>{draft.model}</b></div><div><span>06</span><code>output</code><b>{draft.outputEntity}.{draft.outputField}</b></div><div><span>07</span><code>trigger</code><b>{draft.triggerType}</b></div></div>
    <div className="compile-summary"><p><span>Estimated context</span><b>1,240 tokens</b></p><p><span>Flow pattern</span><b>{triggerPresentation(draft.triggerType, draft.triggerColumns).pattern}</b></p><p><span>Components</span><b>3 assets</b></p></div>
    <div className="compiled-footer"><Code /><span><b>Valid JSON</b><small>summary-recipe.schema.json</small></span><button type="button">View JSON</button></div>
  </aside>;
}

const paneContent: Record<PaneKind, { title: string; eyebrow: string; value: string; technical: string; options: string[] }> = {
  table: { title: "Primary table", eyebrow: "Summary source", value: "Account · account", technical: "Dataverse table", options: ["Account · account", "Case · incident", "Opportunity · opportunity"] },
  execution: { title: "Summary mode", eyebrow: "Summary pattern", value: "One summary per record", technical: "PerRecord", options: ["One summary per record", "Consolidated summary", "Modified records only"] },
  prompt: { title: "AI Prompt", eyebrow: "Reusable asset", value: "Executive summary prompt", technical: "prompt_account_brief_v3", options: ["Executive summary prompt", "Management brief", "Commercial risk summary"] },
  model: { title: "Model", eyebrow: "Generation", value: "GPT-4.1 mini", technical: "EU Data Zone", options: ["GPT-4.1 mini", "GPT-4.1", "GPT-4o mini"] },
  fields: { title: "Selected fields", eyebrow: "Summary context", value: "Account fields", technical: "Dataverse metadata", options: [] },
  records: { title: "Records to process", eyebrow: "Record selection", value: "Custom FetchXML", technical: "csp_fetchxml", options: [] },
  relationships: { title: "Related records", eyebrow: "Context source", value: "Open and recent activities", technical: "csp_relationships", options: ["Open and recent activities", "No related records"] },
  contextQuery: { title: "Context query", eyebrow: "Per-record context", value: "FetchXML per record", technical: "csp_relatedfetchxml", options: [] },
  outputTable: { title: "Destination table", eyebrow: "Summary destination", value: "Account · account", technical: "Dataverse table", options: ["Account · account", "Summary cache · csp_aisummarycache"] },
  outputField: { title: "Destination column", eyebrow: "Summary destination", value: "AI summary · csp_aisummary", technical: "Dataverse column", options: ["AI summary · csp_aisummary", "Description · description", "Name · name"] },
};

const accountFieldOptions = [
  ["name", "Name"], ["accountnumber", "Account number"], ["industrycode", "Industry"],
  ["revenue", "Revenue"], ["description", "Description"], ["primarycontactid", "Primary contact"],
  ["modifiedon", "Last modified"],
] as const;

function FieldsPane({ selectedFields, onClose, onApply }: { selectedFields: string[]; onClose: () => void; onApply: (fields: string[]) => void }) {
  const [selected, setSelected] = useState(selectedFields);
  const toggle = (field: string) => setSelected((current) => current.includes(field) ? current.filter((item) => item !== field) : [...current, field]);
  return <><button className="pane-scrim" type="button" aria-label="Close edit pane" onClick={onClose} /><aside className="edit-pane" aria-label="Edit selected fields"><header><div><p className="micro-label text-brand-blue">Summary context</p><h2>Selected fields</h2></div><button type="button" aria-label="Close pane" onClick={onClose}>×</button></header><div className="pane-current"><span><Database /></span><div><small>Current selection</small><b>{selected.length} Account fields</b><code>account</code></div></div><section><p className="field-caption">Available fields</p>{accountFieldOptions.map(([field, label]) => { const active = selected.includes(field); return <button type="button" key={field} aria-label={`${label} · ${field}`} aria-pressed={active} onClick={() => toggle(field)} className={active ? "selected" : ""}><span>{active ? <Check /> : null}</span><div><b>{label}</b><code>{field}</code></div></button>; })}</section><footer><Button variant="outline" onClick={onClose}>Cancel</Button><Button onClick={() => onApply(selected)} disabled={selected.length === 0}>Apply changes</Button></footer></aside></>;
}

type QueryDesignerMode = "filters" | "view" | "fetchxml";

function relatedSelectionFetchXml(relationship?: Record<string, unknown>) {
  if (relationship?.fetchXml) return String(relationship.fetchXml);
  const days = Number(relationship?.windowDays ?? 30);
  return `<fetch top="12">\n  <entity name="activitypointer">\n    <all-attributes />\n    <filter type="and">\n      <condition attribute="createdon" operator="last-x-days" value="${days}" />\n    </filter>\n    <order attribute="createdon" descending="true" />\n  </entity>\n</fetch>`;
}

function QueryDesignerPane({ purpose, targetEntity, fetchXml, maxRecords, onClose, onApply }: { purpose: "records" | "related"; targetEntity: string; fetchXml: string; maxRecords: number; onClose: () => void; onApply: (value: { fetchXml: string; maxRecords: number; queryMode: number; viewId?: string; viewType?: "system" | "personal" }) => void }) {
  const [mode, setMode] = useState<QueryDesignerMode>("filters");
  const [query, setQuery] = useState(fetchXml);
  const [limit, setLimit] = useState(maxRecords);
  const [viewId, setViewId] = useState("");
  const [viewType, setViewType] = useState<"system" | "personal">("system");
  const systemViews = useSystemViews(targetEntity);
  const userViews = useUserViews(targetEntity);
  const availableViews = viewType === "system" ? (systemViews.data ?? []) : (userViews.data ?? []);
  const title = purpose === "records" ? "Record selection" : "Related records";
  const technical = purpose === "records" ? "csp_fetchxml" : "csp_relatedfetchxml";
  const chooseView = (id: string) => {
    setViewId(id);
    const selected = availableViews.find((view) => view.id === id);
    if (selected?.fetchXml) setQuery(selected.fetchXml);
  };
  const apply = () => onApply({ fetchXml: query.trim(), maxRecords: limit, queryMode: mode === "view" ? (viewType === "system" ? 100000000 : 100000001) : 100000002, viewId: mode === "view" ? viewId : undefined, viewType: mode === "view" ? viewType : undefined });

  return <><button className="pane-scrim" type="button" aria-label="Close edit pane" onClick={onClose} /><aside className="edit-pane query-designer-pane" aria-label={`Configure ${title.toLowerCase()}`}><header><div><p className="micro-label text-brand-blue">{purpose === "records" ? "Rows entering the automation" : "Rows enriching each summary"}</p><h2>{title}</h2></div><button type="button" aria-label="Close pane" onClick={onClose}>×</button></header><div className="query-mode-tabs" role="tablist" aria-label={`${title} method`}><button type="button" role="tab" aria-selected={mode === "filters"} onClick={() => setMode("filters")}><Settings />Build filters</button><button type="button" role="tab" aria-selected={mode === "view"} onClick={() => setMode("view")}><Grid />Choose a view</button><button type="button" role="tab" aria-selected={mode === "fetchxml"} onClick={() => setMode("fetchxml")}><Code />Edit FetchXML</button></div><section className="query-designer-body">
    {mode === "filters" && <><div className="query-mode-intro"><b>Build it like a Dataverse view</b><small>Add conditions, sorting, and a record limit. The app compiles valid FetchXML for the flow.</small></div><VisualFilterBuilder targetEntity={targetEntity} onGenerate={(generated) => setQuery(generated)} /></>}
    {mode === "view" && <><div className="query-mode-intro"><b>Reuse an existing Dataverse view</b><small>Select a system or personal view. Its FetchXML becomes part of this configuration.</small></div><div className="view-type-toggle"><button type="button" className={viewType === "system" ? "active" : ""} onClick={() => { setViewType("system"); setViewId(""); }}>System views</button><button type="button" className={viewType === "personal" ? "active" : ""} onClick={() => { setViewType("personal"); setViewId(""); }}>Personal views</button></div><label className="studio-field"><span>Saved view</span><select aria-label="Saved view" value={viewId} onChange={(event) => chooseView(event.target.value)}><option value="">Select a {viewType} view…</option>{availableViews.map((view) => <option key={view.id} value={view.id}>{view.name}</option>)}</select></label><p className="query-source-status">{(viewType === "system" ? systemViews.isFetching : userViews.isFetching) ? "Loading Dataverse views…" : `${availableViews.length} ${viewType} views available for ${targetEntity}`}</p></>}
    {mode === "fetchxml" && <><div className="query-mode-intro"><b>Advanced FetchXML</b><small>Edit the compiled query directly for joins, nested groups, or operators not exposed by the filter builder.</small></div><label className="studio-field"><span>{purpose === "records" ? "Record selection FetchXML" : "Related-record FetchXML"}</span><textarea aria-label={purpose === "records" ? "Record selection FetchXML" : "Related-record FetchXML"} className="query-code-editor" value={query} onChange={(event) => setQuery(event.target.value)} spellCheck={false} /></label></>}
    <div className="query-output"><span><Code /></span><div><small>Output stored in</small><b>{technical}</b></div><code>{query.length.toLocaleString("en-US")} characters</code></div>
  </section><footer><div className="query-limit"><label>Maximum records <input type="number" min={1} max={5000} value={limit} onChange={(event) => setLimit(Math.max(1, Math.min(5000, Number(event.target.value) || 1)))} /></label></div><Button variant="outline" onClick={onClose}>Cancel</Button><Button onClick={apply} disabled={!query.trim().startsWith("<fetch") || (mode === "view" && !viewId)}>Apply selection</Button></footer></aside></>;
}

function ContextQueryPane({ fetchXml, onClose, onApply }: { fetchXml: string; onClose: () => void; onApply: (fetchXml: string) => void }) {
  const [query, setQuery] = useState(fetchXml);
  return <><button className="pane-scrim" type="button" aria-label="Close edit pane" onClick={onClose} /><aside className="edit-pane" aria-label="Edit context query"><header><div><p className="micro-label text-brand-blue">Per-record context</p><h2>Context FetchXML</h2></div><button type="button" aria-label="Close pane" onClick={onClose}>×</button></header><div className="pane-current"><span><Code /></span><div><small>Configuration column</small><b>Context query</b><code>csp_relatedfetchxml</code></div></div><section className="space-y-4"><div className="query-guidance"><GitBranch /><p><b>Runs once for every selected record</b><small>Use <code>{"{{recordId}}"}</code> where the generated flow must inject the current Dataverse row identifier.</small></p></div><label className="studio-field"><span>Context FetchXML</span><textarea aria-label="Context FetchXML" className="min-h-80 w-full resize-none rounded-lg border border-border bg-slate-950 p-3 font-mono text-xs leading-5 text-cyan-100 outline-none focus:border-brand-cyan" value={query} onChange={(event) => setQuery(event.target.value)} spellCheck={false} /></label><p className="text-xs leading-5 text-muted-foreground">This query builds the payload sent to the AI Prompt. It can include source fields, linked entities, filters, ordering, and record limits.</p></section><footer><Button variant="outline" onClick={onClose}>Cancel</Button><Button onClick={() => onApply(query.trim())} disabled={!query.trim().startsWith("<fetch") || !query.includes("{{recordId}}")}>Apply changes</Button></footer></aside></>;
}

function EditPane({ pane, current, options, onClose, onApply }: { pane: PaneKind; current: string; options: string[]; onClose: () => void; onApply: (value: string) => void }) {
  const content = paneContent[pane];
  const [selected, setSelected] = useState(current);
  return <><button className="pane-scrim" type="button" aria-label="Close edit pane" onClick={onClose} /><aside className="edit-pane" aria-label={`Edit ${content.title.toLowerCase()}`}><header><div><p className="micro-label text-brand-blue">{content.eyebrow}</p><h2>{content.title}</h2></div><button type="button" aria-label="Close pane" onClick={onClose}>×</button></header><div className="pane-current"><span><CheckCircle /></span><div><small>Current value</small><b>{current}</b><code>{content.technical}</code></div></div><section><p className="field-caption">Select value</p>{options.map((option) => <button type="button" key={option} onClick={() => setSelected(option)} className={selected === option ? "selected" : ""}><span>{selected === option ? <Check /> : null}</span>{option}</button>)}</section><footer><Button variant="outline" onClick={onClose}>Cancel</Button><Button onClick={() => onApply(selected)}>Apply changes</Button></footer></aside></>;
}

function Builder() {
  const { data } = useStudioRuntime();
  const { configurationId } = useParams();
  const isNew = !configurationId;
  const createConfiguration = useCreateStudioConfiguration();
  const updateConfiguration = useUpdateStudioConfiguration();
  const updatePrompt = useUpdateStudioPrompt();
  const configuration = configurationId === "account-operations"
    ? data?.configurations[0]
    : data?.configurations.find((item) => item.id === configurationId);
  const prompts = useMemo(() => data?.prompts ?? [], [data?.prompts]);
  const [step, setStep] = useState(0);
  const [recipeOpen, setRecipeOpen] = useState(false);
  const [pane, setPane] = useState<PaneKind | null>(null);
  const [validationErrors, setValidationErrors] = useState<string[]>([]);
  const [draft, setDraft] = useState<BuilderDraft>(() => draftFromConfiguration());
  const navigate = useNavigate();
  // Dataverse arrives asynchronously after the local-first editor has mounted.
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect
    if (configuration) setDraft(draftFromConfiguration(configuration, prompts));
    else if (isNew && prompts.length) setDraft((current) => current.promptId ? current : draftFromConfiguration(undefined, prompts));
  }, [configuration, isNew, prompts]);
  const save = async (next: BuilderDraft, changes: Record<string, unknown>) => {
    setDraft(next);
    setPane(null);
    setValidationErrors([]);
    if (configuration) {
      await updateConfiguration.mutateAsync({ id: configuration.id, changes: { ...changes, csp_configurationjson: JSON.stringify(compileRecipe(next)) } });
      return configuration.id;
    }
    if (isNew) {
      const id = await createConfiguration.mutateAsync(buildDraftConfigurationPayload(next));
      navigate(`/configurations/${id}`, { replace: true });
      return id;
    }
    return undefined;
  };
  const paneOptions = pane === "prompt" ? prompts.map((item) => item.name) : pane ? paneContent[pane].options : [];
  const paneCurrent = pane === "table" ? ({ account: "Account · account", incident: "Case · incident", opportunity: "Opportunity · opportunity" }[draft.entity] ?? draft.entity)
    : pane === "execution" ? (draft.mode === 100000000 ? "One summary per record" : "Consolidated summary")
    : pane === "prompt" ? draft.promptName
    : pane === "outputTable" ? (draft.outputEntity === "account" ? "Account · account" : "Summary cache · csp_aisummarycache")
    : pane === "outputField" ? ({ csp_aisummary: "AI summary · csp_aisummary", description: "Description · description", name: "Name · name" }[draft.outputField] ?? draft.outputField)
    : pane === "records" ? "Custom FetchXML"
    : pane === "relationships" ? (draft.relationships.length ? "Open and recent activities" : "No related records")
    : pane === "contextQuery" ? "FetchXML per record"
    : draft.model.replace("gpt-", "GPT-").replace("-mini", " mini");
  const applyPane = async (value: string) => {
    if (!pane) return;
    if (pane === "table") { const entity = value.includes("incident") ? "incident" : value.includes("opportunity") ? "opportunity" : "account"; const fetchXml = buildRecordSelectionFetchXml(entity, draft.maxRecords); const relatedFetchXml = buildContextFetchXml(entity, draft.sourceFields, draft.relationships); const target = getTargetEntityIdentity(entity); await save({ ...draft, entity, fetchXml, relatedFetchXml }, { csp_targetentity: entity, csp_targetentityset: target.entitySet, csp_targetentityidfield: target.idField, csp_fetchxml: fetchXml, csp_relatedfetchxml: relatedFetchXml }); }
    if (pane === "execution") { const mode = value === "Consolidated summary" ? 100000001 : 100000000; await save({ ...draft, mode }, { csp_mode: mode }); }
    if (pane === "model") { const model = value.toLowerCase().replace(" ", "-"); await save({ ...draft, model }, { csp_model: model }); }
    if (pane === "prompt") { const prompt = prompts.find((item) => item.name === value); if (prompt) await save({ ...draft, promptId: prompt.id, promptName: prompt.name, promptKey: prompt.key, promptContent: prompt.content }, { "csp_Prompt@odata.bind": `/csp_aiprompts(${prompt.id})` }); }
    if (pane === "outputTable") { const outputEntity = value.includes("csp_aisummarycache") ? "csp_aisummarycache" : "account"; await save({ ...draft, outputEntity }, { csp_outputentity: outputEntity }); }
    if (pane === "outputField") { const outputField = value.includes("description") ? "description" : value.includes(" · name") ? "name" : "csp_aisummary"; await save({ ...draft, outputField }, { csp_outputfield: outputField }); }
  };
  const updateDraft = async (changes: Partial<BuilderDraft>) => {
    const next = { ...draft, ...changes };
    const fields: Record<string, unknown> = {};
    if (changes.sourceFields || changes.relationships) {
      next.relatedFetchXml = buildContextFetchXml(next.entity, next.sourceFields, next.relationships);
      fields.csp_sourcefields = JSON.stringify(next.sourceFields);
      fields.csp_relationships = JSON.stringify(next.relationships);
      fields.csp_relatedfetchxml = next.relatedFetchXml;
    }
    if (changes.triggerColumns) fields.csp_triggercolumns = JSON.stringify(changes.triggerColumns);
    if (changes.outputEntity) fields.csp_outputentity = changes.outputEntity;
    if (changes.outputField) fields.csp_outputfield = changes.outputField;
    await save(next, fields);
  };
  const savePromptContent = async (content: string) => {
    if (!draft.promptId) return;
    await updatePrompt.mutateAsync({ id: draft.promptId, content });
    setDraft((current) => ({ ...current, promptContent: content }));
  };
  const publish = async () => {
    const errors = validateSummaryConfiguration(draft);
    setValidationErrors(errors);
    if (errors.length) {
      setStep(4);
      return;
    }
    let id = configuration?.id;
    if (configuration) {
      await updateConfiguration.mutateAsync({ id: configuration.id, changes: buildConfigurationPayload(draft) });
    } else if (isNew) {
      id = await createConfiguration.mutateAsync(buildConfigurationPayload(draft));
    }
    if (!id) return;
    await updateConfiguration.mutateAsync({ id, changes: buildPublishSignal() });
    navigate(`/published/${id}`);
  };
  const content = [<DataStep onEdit={setPane} draft={draft} accounts={data?.accounts ?? []} />, <PromptStep key={draft.promptId} onEdit={setPane} draft={draft} onSavePrompt={savePromptContent} savingPrompt={updatePrompt.isPending} />, <DestinationStep draft={draft} onEdit={setPane} onChange={updateDraft} />, <ExecutionStep draft={draft} onChange={updateDraft} />, <ReviewStep draft={draft} validationErrors={validationErrors} />][step];
  return <Shell><div className="studio-page builder-page">
    <PageHeader eyebrow={`Configuration · ${configuration?.version ?? "draft"}`} title={draft.name} description="" actions={<><span className="saved-state"><CheckCircle />{updateConfiguration.isPending || createConfiguration.isPending ? "Saving…" : data?.live ? "Dataverse" : "Local demo"}</span><Button type="button" variant="outline" size="sm" onClick={() => setRecipeOpen(!recipeOpen)}><Code data-icon="inline-start" />{recipeOpen ? "Hide recipe" : "View recipe"}</Button><Button size="sm">Test with a record</Button></>} />
    {recipeOpen && <div className="recipe-drawer"><CompiledRecipe step={step} draft={draft} title="Compiled recipe" /></div>}
    <div className="builder-layout">
      <aside className="step-rail bpf-process" aria-label="Configuration process"><div className="bpf-process-header"><span>Configuration process</span><em>Stage {step + 1} of {steps.length}</em></div><div className="bpf-stages">{steps.map((item, index) => { const Icon = item.icon; return <button type="button" key={item.title} onClick={() => setStep(index)} className={cn(step === index && "active", index < step && "complete")} aria-current={step === index ? "step" : undefined}><span>{index < step ? <Check /> : <Icon />}</span><div><small>0{index+1} · {item.stage}</small><b>{item.title}</b></div></button>; })}</div></aside>
      <main className="editor-panel">{content}<footer className="editor-footer" aria-label="Process actions"><Button variant="ghost" disabled={step === 0} onClick={() => setStep(step - 1)}>Previous</Button><span>Step {step + 1} of {steps.length}</span>{step < 4 ? <Button onClick={() => setStep(step + 1)}>Continue <ArrowRight data-icon="inline-end" /></Button> : <Button onClick={publish} disabled={updateConfiguration.isPending || createConfiguration.isPending}><Flow data-icon="inline-start" />{updateConfiguration.isPending || createConfiguration.isPending ? "Saving…" : "Publish configuration"}</Button>}</footer></main>
    </div>
    {pane === "fields" ? <FieldsPane selectedFields={draft.sourceFields} onClose={() => setPane(null)} onApply={(fields) => updateDraft({ sourceFields: fields })} /> : pane === "records" ? <QueryDesignerPane purpose="records" targetEntity={draft.entity} fetchXml={draft.fetchXml} maxRecords={draft.maxRecords} onClose={() => setPane(null)} onApply={({ fetchXml, maxRecords, queryMode, viewId, viewType }) => save({ ...draft, fetchXml, maxRecords, queryMode }, { csp_fetchxml: fetchXml, csp_maxrecords: maxRecords, csp_querymode: queryMode, ...(viewId && viewType === "system" ? { csp_systemviewid: viewId } : {}), ...(viewId && viewType === "personal" ? { csp_userviewid: viewId } : {}) })} /> : pane === "relationships" ? <QueryDesignerPane purpose="related" targetEntity={String(draft.relationships[0]?.entity ?? "activitypointer")} fetchXml={relatedSelectionFetchXml(draft.relationships[0])} maxRecords={Number(draft.relationships[0]?.maxRecords ?? 12)} onClose={() => setPane(null)} onApply={({ fetchXml, maxRecords, queryMode, viewId, viewType }) => { const relationships = [{ ...(draft.relationships[0] ?? {}), entity: String(draft.relationships[0]?.entity ?? "activitypointer"), maxRecords, fetchXml, selectionMode: queryMode, ...(viewId ? { viewId, viewType } : {}) }]; const relatedFetchXml = buildContextFetchXml(draft.entity, draft.sourceFields, relationships); return save({ ...draft, relationships, relatedFetchXml }, { csp_relationships: JSON.stringify(relationships), csp_relatedfetchxml: relatedFetchXml }); }} /> : pane === "contextQuery" ? <ContextQueryPane fetchXml={draft.relatedFetchXml} onClose={() => setPane(null)} onApply={(relatedFetchXml) => save({ ...draft, relatedFetchXml }, { csp_relatedfetchxml: relatedFetchXml })} /> : pane && <EditPane pane={pane} current={paneCurrent} options={paneOptions.length ? paneOptions : paneContent[pane].options} onClose={() => setPane(null)} onApply={applyPane} />}
  </div></Shell>;
}

function Published() {
  return <Shell><div className="studio-page publish-page">
    <div className="success-hero"><div className="success-mark"><Check /></div><p className="micro-label text-brand-cyan">VERSIONED CONFIGURATION IN DATAVERSE</p><h1>Recipe prepared</h1><p>The configuration is ready for the backend to generate or update the specialized Power Automate flow.</p><div className="mt-6 flex justify-center gap-3"><Button><Play data-icon="inline-start" />Test configuration</Button><Button variant="outline" className="border-white/20 bg-white/5 text-white hover:bg-white/10 hover:text-white">View JSON contract <ArrowRight data-icon="inline-end" /></Button></div><div className="deployment-checks"><span><CheckCircle />Schema validated</span><span><LinkIcon />AI Prompt bound</span><span><Blocks />Stored in the solution</span></div></div>
    <section className="publication-assets"><div className="publication-title"><div><p className="micro-label text-muted-foreground">Compilation result</p><h2>3 assets prepared</h2></div><span><i />Demo · Irish Power Platform Summit 2026</span></div>
      <div className="asset-timeline"><PublishedAsset icon={<Database />} title="Configuration stored" technical="csp_aisummaryconfig · v3.0" detail="Versioned recipe available for audit." time="Ready" /><PublishedAsset icon={<Flow />} title="Flow definition prepared" technical="csp_SUM_AccountOperations_v3" detail="Trigger, query, prompt, and write contract for the backend." time="Pending" featured /><PublishedAsset icon={<Bot />} title="AI Prompt bound" technical="account-operations" detail="Inputs mapped and GPT-4.1 mini validated." time="Ready" /></div>
    </section>
    <section className="execution-ready"><div><p className="micro-label text-brand-blue">Next step · backend</p><h2>Contract ready to materialize</h2><p>The backend reads the recipe, composes the context, and updates <code>account.csp_aisummary</code>; the Code App already uses the same Dataverse model.</p></div><div className="ready-flow"><span><Database /><small>Recipe</small></span><i /><span><Flow /><small>Backend</small></span><i /><span><Bot /><small>AI Prompt</small></span><i /><span><Target /><small>Summary</small></span></div><Button asChild variant="outline"><Link to="/">Back to configurations</Link></Button></section>
  </div></Shell>;
}

function PublishedAsset({ icon, title, technical, detail, time, featured }: { icon: ReactNode; title: string; technical: string; detail: string; time: string; featured?: boolean }) { return <div className={cn("published-asset", featured && "featured")}><span className="asset-check"><Check /></span><div className="asset-symbol">{icon}</div><div><p>{title}</p><code>{technical}</code><small>{detail}</small></div><em>{time}</em></div>; }

function Runs() {
  const { data, refetch, isFetching } = useStudioRuntime();
  const runs = data?.runs ?? [];
  const successes = runs.filter((run) => run.status === 100000000).length;
  const average = runs.length ? runs.reduce((total, run) => total + run.latencyMs, 0) / runs.length : 0;
  const tokens = runs.reduce((total, run) => total + run.tokens, 0);
  const formatDate = (value: string | null) => value ? new Intl.DateTimeFormat("en-IE", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }).format(new Date(value)) : "Not available";
  return <Shell><div className="studio-page"><PageHeader eyebrow="Operations" title="Runs" description="Inspect which recipe ran, which record it processed, and what result was written to Dataverse." actions={<Button variant="outline" onClick={() => refetch()} disabled={isFetching}><RefreshCw data-icon="inline-start" />{isFetching ? "Refreshing" : "Refresh"}</Button>} /><div className="run-summary"><div><ActivityStat value={String(runs.length)} label="Recorded" /><ActivityStat value={`${runs.length ? ((successes / runs.length) * 100).toFixed(1) : "0"}%`} label="Successful" /><ActivityStat value={`${(average / 1000).toFixed(1)} s`} label="Average duration" /><ActivityStat value={tokens.toLocaleString("en-IE")} label="Tokens" /></div><span><i />{data?.live ? "Dataverse connected" : "Local demo"}</span></div><div className="runs-table"><div className="runs-head"><span>Run</span><span>Configuration</span><span>Record</span><span>Status</span><span>Duration</span><span>Started</span></div>{runs.map((run) => { const status = run.status === 100000000 ? "Successful" : run.status === 100000001 ? "Running" : "Error"; return <div className="runs-row" key={run.id}><span className="mono">{run.name}</span><span>{data?.configurations[0]?.name ?? "Summary"}</span><span>{run.accountName}</span><span className={status === "Successful" ? "run-ok" : "run-warn"}>{status}</span><span>{(run.latencyMs / 1000).toFixed(1)} s</span><span>{formatDate(run.timestamp)}</span></div>; })}</div></div></Shell>;
}
function ActivityStat({ value, label }: { value: string; label: string }) { return <span><b>{value}</b><small>{label}</small></span>; }

export default function App() {
  return <Routes><Route path="/" element={<Overview />} /><Route path="/configurations/new" element={<Builder />} /><Route path="/configurations/:configurationId" element={<Builder />} /><Route path="/published/:configurationId" element={<Published />} /><Route path="/runs" element={<Runs />} /></Routes>;
}
