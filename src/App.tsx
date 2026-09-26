import {
  Fragment,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
} from "react";
import {
  Avatar,
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  Menu,
  MenuItem,
  MenuItemLink,
  MenuList,
  MenuPopover,
  MenuTrigger,
  NavDrawer,
  NavDrawerBody,
  NavDrawerHeader,
  NavItem,
  NavSectionHeader,
  SearchBox,
  Select,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Textarea,
  Toast,
  ToastBody,
  ToastTitle,
  Toaster,
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
  useToastController,
} from "@fluentui/react-components";
import { Apps20Regular as Blocks } from "@fluentui/react-icons/svg/apps";
import { ArrowRight20Regular as ArrowRight } from "@fluentui/react-icons/svg/arrow-right";
import { ArrowSync20Regular as RefreshCw } from "@fluentui/react-icons/svg/arrow-sync";
import { Bot20Regular as Bot } from "@fluentui/react-icons/svg/bot";
import { Branch20Regular as GitBranch } from "@fluentui/react-icons/svg/branch";
import {
  Braces20Filled,
  Braces20Regular as Braces,
} from "@fluentui/react-icons/svg/braces";
import { Checkmark20Regular as Check } from "@fluentui/react-icons/svg/checkmark";
import { CheckmarkCircle20Regular as CheckCircle } from "@fluentui/react-icons/svg/checkmark-circle";
import { ChevronRight20Regular as ChevronRight } from "@fluentui/react-icons/svg/chevron-right";
import { Clock20Regular as Clock } from "@fluentui/react-icons/svg/clock";
import { Copy20Regular as Copy } from "@fluentui/react-icons/svg/copy";
import { Delete20Regular as Delete } from "@fluentui/react-icons/svg/delete";
import {
  DocumentText20Filled,
  DocumentText20Regular as DocumentText,
} from "@fluentui/react-icons/svg/document-text";
import { ErrorCircle20Regular as ErrorCircle } from "@fluentui/react-icons/svg/error-circle";
import { MoreHorizontal20Regular as More } from "@fluentui/react-icons/svg/more-horizontal";
import { Open20Regular as OpenIcon } from "@fluentui/react-icons/svg/open";
import { Warning20Regular as Warning } from "@fluentui/react-icons/svg/warning";
import { Database20Regular as Database } from "@fluentui/react-icons/svg/database";
import { Dismiss20Regular as Dismiss } from "@fluentui/react-icons/svg/dismiss";
import { Person20Regular as Person } from "@fluentui/react-icons/svg/person";
import { Flow20Regular as Flow } from "@fluentui/react-icons/svg/flow";
import {
  Grid20Filled,
  Grid20Regular as Grid,
} from "@fluentui/react-icons/svg/grid";
import { History20Regular as History } from "@fluentui/react-icons/svg/history";
import { Link20Regular as LinkIcon } from "@fluentui/react-icons/svg/link";
import { Play20Regular as Play } from "@fluentui/react-icons/svg/play";
import { Search20Regular as Search } from "@fluentui/react-icons/svg/search";
import { Settings20Regular as Settings } from "@fluentui/react-icons/svg/settings";
import { Sparkle20Regular as Sparkles } from "@fluentui/react-icons/svg/sparkle";
import { Target20Regular as Target } from "@fluentui/react-icons/svg/target";
import { WindowConsole20Regular as Code } from "@fluentui/react-icons/svg/window-console";
import {
  Link,
  Route,
  Routes,
  useLocation,
  useNavigate,
  useParams,
} from "react-router-dom";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { VisualFilterBuilder } from "@/components/VisualFilterBuilder";
import { cn } from "@/lib/utils";
import {
  buildConfigurationPayload,
  buildContextFetchXml,
  buildDraftConfigurationPayload,
  buildPublishSignal,
  buildRecordSelectionFetchXml,
  compileRecipe,
  estimateRecipeTokens,
  getTargetEntityIdentity,
  validateSummaryConfiguration,
  type SummaryConfigurationDraft,
} from "@/lib/summaryConfiguration";
import { buildPromptDraft, type PromptWizardAnswers } from "@/lib/promptWizard";
import { formatFetchXml } from "@/lib/fetchxml";
import brandLogo from "@/assets/creativity-spark-mark.png?inline";
import {
  STUDIO_SOLUTION_UNIQUE_NAME,
  cloudFlowRunUrl,
  cloudFlowUrl,
  relatedCloudFlows,
  solutionUrl,
  useCloudFlows,
  useStudioSolution,
} from "@/hooks/useCloudFlows";
import { useOutputSummaries } from "@/hooks/useOutputSummaries";
import {
  useCreateStudioConfiguration,
  useCreateStudioPrompt,
  useDeactivateStudioConfiguration,
  useDeleteStudioConfiguration,
  useSaveStudioPrompt,
  useStudioRuntime,
  useUpdateStudioConfiguration,
  useUpdateStudioPrompt,
  type StudioAccount,
  type StudioConfiguration,
  type StudioPrompt,
  type StudioPromptInput,
  type StudioRun,
  type StudioRuntime,
} from "@/hooks/useStudioRuntime";
import { useSystemViews } from "@/hooks/useSystemViews";
import { useUserViews } from "@/hooks/useUserViews";
import {
  useDataverseColumns,
  useDataverseRecordPreview,
  useDataverseRelationships,
  useDataverseTableMetadata,
  useDataverseTables,
} from "@/hooks/useDataverseCatalog";
import type {
  DataverseColumnMetadata,
  DataverseRelationshipMetadata,
  DataverseTableMetadata,
} from "@/lib/summaryConfiguration";

function BrandMark({ className }: { className?: string }) {
  return (
    <svg
      className={className}
      viewBox="0 0 64 94"
      fill="none"
      aria-hidden="true"
      focusable="false"
    >
      <defs>
        <linearGradient
          id="spark-brand-gradient"
          x1="12"
          y1="8"
          x2="54"
          y2="74"
          gradientUnits="userSpaceOnUse"
        >
          <stop stopColor="#02C8DD" />
          <stop offset="1" stopColor="#0085C8" />
        </linearGradient>
      </defs>
      <path
        d="M35.8 4.5c8.8 7.7 10.6 17.1 2.2 24.8L20.4 45.5c-4.2 3.9-4 8.2.7 12.3l7.7 6.8-8.4 8.1-8.5-7.5C1.4 55.9 1.2 42.9 11 33.9l16.8-15.4c4.5-4.1 5.5-8.5 2-12.6L35.8 0v4.5Z"
        fill="url(#spark-brand-gradient)"
      />
      <path
        d="m43.7 20.1 8.4 7.4c10.6 9.3 10.8 22.2 1 31.2L36.3 74.1c-4.5 4.1-5.4 8.4-2 12.6l-6 5.9v-4.5c-8.8-7.7-10.6-17.1-2.2-24.8l17.6-16.2c4.2-3.9 4-8.2-.7-12.3l-7.7-6.8 8.4-7.9Z"
        fill="url(#spark-brand-gradient)"
      />
    </svg>
  );
}

const studioToasterId = "summary-studio-toaster";
type ToastIntent = "success" | "error" | "info" | "warning";

function useStudioToast() {
  const { dispatchToast } = useToastController(studioToasterId);
  return useCallback(
    (intent: ToastIntent, title: string, body?: string) => {
      dispatchToast(
        <Toast>
          <ToastTitle>{title}</ToastTitle>
          {body ? <ToastBody>{body}</ToastBody> : null}
        </Toast>,
        { intent, timeout: intent === "error" ? 8000 : 4000 },
      );
    },
    [dispatchToast],
  );
}

function errorMessage(error: unknown, fallback: string) {
  return error instanceof Error && error.message ? error.message : fallback;
}

function BrandLogo() {
  const [imageFailed, setImageFailed] = useState(false);
  if (imageFailed) return <BrandMark />;
  return (
    <img
      src={brandLogo}
      alt=""
      decoding="async"
      onError={() => setImageFailed(true)}
    />
  );
}

function Shell({ children }: { children: ReactNode }) {
  const { data, refetch, isFetching } = useStudioRuntime();
  const live = data?.live ?? false;
  const user = data?.user ?? null;
  const navigate = useNavigate();
  const location = useLocation();
  const [navigationCollapsed, setNavigationCollapsed] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const environmentId = data?.environmentId ?? "";
  const makerHome = environmentId
    ? `https://make.powerapps.com/environments/${environmentId}/home`
    : "https://make.powerapps.com";
  const automateHome = environmentId
    ? `https://make.powerautomate.com/environments/${environmentId}/flows`
    : "https://make.powerautomate.com";
  const tablesHome = environmentId
    ? `https://make.powerapps.com/environments/${environmentId}/tables`
    : "https://make.powerapps.com";
  const solution = useStudioSolution();
  const solutionHome = solutionUrl(environmentId, solution.data?.id ?? "");
  const solutionLabel = solution.data?.friendlyName ?? STUDIO_SOLUTION_UNIQUE_NAME;
  const searchValue = new URLSearchParams(location.search).get("q") ?? "";
  const firstConfigurationId = data?.configurations[0]?.id;
  const designPath = firstConfigurationId
    ? `/configurations/${firstConfigurationId}`
    : "/configurations/new";
  const navItems = [
    { to: "/", label: "Configurations", icon: Grid, activeIcon: Grid20Filled },
    {
      to: designPath,
      label: "Design summary",
      icon: Braces,
      activeIcon: Braces20Filled,
    },
    { to: "/prompts", label: "Prompts", icon: Bot, activeIcon: Bot },
    {
      to: "/summaries",
      label: "Summaries",
      icon: DocumentText,
      activeIcon: DocumentText20Filled,
    },
    { to: "/runs", label: "Runs", icon: History, activeIcon: History },
  ];
  const followInternalLink = (
    event: ReactMouseEvent<HTMLElement>,
    to: string,
  ) => {
    if (
      event.defaultPrevented ||
      event.button !== 0 ||
      event.metaKey ||
      event.ctrlKey ||
      event.shiftKey ||
      event.altKey
    ) {
      return;
    }
    event.preventDefault();
    navigate(to);
  };
  const selectedNavigation = location.pathname.startsWith("/runs")
    ? "/runs"
    : location.pathname.startsWith("/prompts")
      ? "/prompts"
    : location.pathname.startsWith("/summaries")
      ? "/summaries"
    : location.pathname.startsWith("/configurations") ||
        location.pathname.startsWith("/published")
      ? designPath
      : "/";
  return (
    <div className="studio-shell">
      <a href="#main-content" className="skip-link">
        Skip to main content
      </a>
      <aside className="studio-sidebar" data-collapsed={navigationCollapsed}>
        <NavDrawer
          className="studio-nav-drawer"
          open
          type="inline"
          density="medium"
          selectedValue={selectedNavigation}
          aria-label="Summary Studio navigation"
        >
          <NavDrawerHeader className="studio-nav-header">
            <div
              role="img"
              aria-label="Creativity Spark"
              className="brand-lockup"
            >
              <BrandLogo />
              <div className="brand-copy">
                <p>Summary Studio</p>
                <p>Power Platform</p>
              </div>
            </div>
          </NavDrawerHeader>
          <NavDrawerBody>
            <NavSectionHeader>Workspace</NavSectionHeader>
          {navItems.map(
            ({ to, label, icon: Icon, activeIcon: ActiveIcon }) => (
              <NavItem
                key={to}
                value={to}
                href={`#${to}`}
                icon={
                  selectedNavigation === to ? <ActiveIcon /> : <Icon />
                }
                aria-label={label}
                title={navigationCollapsed ? label : undefined}
                onClick={(event) => followInternalLink(event, to)}
              >
                <span className="nav-label">{label}</span>
              </NavItem>
            ),
          )}
          </NavDrawerBody>
        </NavDrawer>
        <button
          className="nav-collapse"
          type="button"
          aria-label={
            navigationCollapsed ? "Expand navigation" : "Collapse navigation"
          }
          onClick={() => setNavigationCollapsed((value) => !value)}
        >
          <ChevronRight />
        </button>
      </aside>
      <main id="main-content" className="min-w-0 flex-1">
        <header className="studio-topbar">
          <div className="studio-topbar-leading">
            <Menu positioning="below-start">
              <MenuTrigger disableButtonEnhancement>
                <Button
                  variant="ghost"
                  size="icon-sm"
                  aria-label="Microsoft 365 apps"
                  title="Power Platform shortcuts"
                >
                  <Blocks />
                </Button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItemLink href={makerHome} target="_blank" rel="noopener">
                    Power Apps
                  </MenuItemLink>
                  <MenuItemLink href={automateHome} target="_blank" rel="noopener">
                    Power Automate
                  </MenuItemLink>
                  <MenuItemLink href={tablesHome} target="_blank" rel="noopener">
                    Dataverse tables
                  </MenuItemLink>
                  {solutionHome && (
                    <MenuItemLink href={solutionHome} target="_blank" rel="noopener">
                      Solution · {solutionLabel}
                    </MenuItemLink>
                  )}
                </MenuList>
              </MenuPopover>
            </Menu>
            <span className="studio-topbar-divider" aria-hidden="true" />
            <Breadcrumb aria-label="Breadcrumb">
              <BreadcrumbItem>
                <BreadcrumbButton
                  href="#/"
                  onClick={(event) => followInternalLink(event, "/")}
                >
                  Summary Studio
                </BreadcrumbButton>
              </BreadcrumbItem>
              <BreadcrumbDivider />
              <BreadcrumbItem>
                <BreadcrumbButton current>
                  {selectedNavigation === "/runs"
                    ? "Runs"
                    : selectedNavigation === "/prompts"
                      ? "Prompts"
                      : selectedNavigation === "/summaries"
                        ? "Summaries"
                        : "Configurations"}
                </BreadcrumbButton>
              </BreadcrumbItem>
            </Breadcrumb>
          </div>
          <SearchBox
            className="studio-global-search"
            aria-label="Search"
            placeholder="Search configurations"
            size="medium"
            value={searchValue}
            onChange={(_, data) => {
              const query = data.value.trim();
              navigate(query ? `/?q=${encodeURIComponent(query)}` : "/", {
                replace: true,
              });
            }}
          />
          <div className="studio-global-actions">
            <span className="environment-status">
              <i />
              {live ? "Dataverse" : "Local"}
            </span>
            <Button
              variant="ghost"
              size="icon-sm"
              aria-label="Settings"
              title="Studio settings"
              aria-expanded={settingsOpen}
              onClick={() => setSettingsOpen(true)}
            >
              <Settings />
            </Button>
            <Avatar
              name={user?.name}
              icon={<Person />}
              size={28}
              color={user ? "colorful" : "neutral"}
              aria-label={user ? user.name : "Signed-in user"}
              title={user?.email || undefined}
            />
          </div>
        </header>
        {children}
      </main>
      <Toaster toasterId={studioToasterId} position="bottom-end" />
      {settingsOpen && (
        <SettingsPane
          runtime={data}
          refreshing={isFetching}
          onRefresh={() => refetch()}
          onClose={() => setSettingsOpen(false)}
          links={{ makerHome, automateHome, tablesHome, solutionHome, solutionLabel }}
        />
      )}
    </div>
  );
}

function SettingsPane({
  runtime,
  refreshing,
  onRefresh,
  onClose,
  links,
}: {
  runtime: StudioRuntime | undefined;
  refreshing: boolean;
  onRefresh: () => void;
  onClose: () => void;
  links: {
    makerHome: string;
    automateHome: string;
    tablesHome: string;
    solutionHome: string;
    solutionLabel: string;
  };
}) {
  const live = runtime?.live ?? false;
  const tables = [
    "csp_aisummaryconfig",
    "csp_aiprompt",
    "csp_aisummarycache",
    "csp_aiusage",
    "account",
    "savedquery",
    "userquery",
  ];
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close settings"
        onClick={onClose}
      />
      <aside className="edit-pane settings-pane" aria-label="Studio settings">
        <header>
          <div>
            <p className="micro-label text-brand-blue">Summary Studio</p>
            <h2>Settings</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <section>
          <p className="field-caption">Connection</p>
          <dl className="settings-list">
            <div>
              <dt>Data source</dt>
              <dd>
                <span className={live ? "status-pill live" : "status-pill"}>
                  <i />
                  {live ? "Dataverse connected" : "Local demo data"}
                </span>
              </dd>
            </div>
            <div>
              <dt>Environment</dt>
              <dd>
                <code>{runtime?.environmentId || "Not available outside Power Apps"}</code>
              </dd>
            </div>
            <div>
              <dt>Organization URL</dt>
              <dd>
                <code>{runtime?.orgUrl || "Not available outside Power Apps"}</code>
              </dd>
            </div>
            <div>
              <dt>Signed in as</dt>
              <dd>
                {runtime?.user ? (
                  <>
                    <b>{runtime.user.name}</b>
                    {runtime.user.email && <small>{runtime.user.email}</small>}
                  </>
                ) : (
                  <small>Anonymous local session</small>
                )}
              </dd>
            </div>
          </dl>

          <p className="field-caption mt-5">Dataverse tables in use</p>
          <div className="configuration-tokens settings-tokens">
            {tables.map((table) => (
              <code key={table}>{table}</code>
            ))}
          </div>

          <p className="field-caption mt-5">Data</p>
          <div className="settings-actions">
            <Button
              type="button"
              variant="outline"
              size="sm"
              disabled={refreshing}
              onClick={onRefresh}
            >
              <RefreshCw data-icon="inline-start" />
              {refreshing ? "Refreshing…" : "Refresh Dataverse data"}
            </Button>
            <small>
              {runtime?.configurations.length ?? 0} configurations ·{" "}
              {runtime?.prompts.length ?? 0} prompts · {runtime?.runs.length ?? 0} runs
              loaded
            </small>
          </div>

          <p className="field-caption mt-5">Open in Power Platform</p>
          <div className="settings-links">
            <a href={links.makerHome} target="_blank" rel="noopener noreferrer">
              <Blocks /> Power Apps maker portal <ArrowRight />
            </a>
            <a href={links.automateHome} target="_blank" rel="noopener noreferrer">
              <Flow /> Power Automate flows <ArrowRight />
            </a>
            <a href={links.tablesHome} target="_blank" rel="noopener noreferrer">
              <Database /> Dataverse tables <ArrowRight />
            </a>
            {links.solutionHome && (
              <a href={links.solutionHome} target="_blank" rel="noopener noreferrer">
                <Blocks /> Solution {links.solutionLabel} (tables, prompts, flows) <ArrowRight />
              </a>
            )}
          </div>

          <p className="field-caption mt-5">About</p>
          <p className="settings-about">
            Summary Studio · Irish Power Platform Summit 2026 demo. Configurations
            are stored as versioned recipes in Dataverse and materialized as
            cloud flows by the generator backend.
          </p>
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Close
          </Button>
        </footer>
      </aside>
    </>
  );
}

function formatRelativeDate(value: string) {
  const elapsed = Date.now() - new Date(value).getTime();
  if (!Number.isFinite(elapsed) || elapsed < 0) return "just now";
  const minutes = Math.round(elapsed / 60_000);
  if (minutes < 1) return "just now";
  if (minutes < 60) return `${minutes} min ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours} h ago`;
  const days = Math.round(hours / 24);
  if (days < 30) return `${days} d ago`;
  return new Intl.DateTimeFormat("en-IE", { day: "2-digit", month: "short" }).format(
    new Date(value),
  );
}

function PageHeader({
  eyebrow,
  title,
  description,
  actions,
}: {
  eyebrow: string;
  title: string;
  description: string;
  actions?: ReactNode;
}) {
  return (
    <div className="page-header">
      <div>
        <p className="micro-label text-brand-blue">{eyebrow}</p>
        <h1>{title}</h1>
        <p>{description}</p>
      </div>
      {actions && (
        <Toolbar
          aria-label={`${title} commands`}
          className="page-command-bar"
        >
          {actions}
        </Toolbar>
      )}
    </div>
  );
}

function Overview() {
  const navigate = useNavigate();
  const location = useLocation();
  const { data, refetch, isFetching } = useStudioRuntime();
  const runtime = data;
  const runtimeConfigurations = runtime?.configurations;
  const configurations = useMemo(
    () => runtimeConfigurations ?? [],
    [runtimeConfigurations],
  );
  const flows = useCloudFlows();
  const notify = useStudioToast();
  const createConfiguration = useCreateStudioConfiguration();
  const deactivateConfiguration = useDeactivateStudioConfiguration();
  const deleteConfiguration = useDeleteStudioConfiguration();
  const [pendingAction, setPendingAction] = useState<{
    kind: "deactivate" | "delete";
    item: StudioConfiguration;
  } | null>(null);
  const generatedFlows = useMemo(() => {
    const seen = new Map<string, { linked: boolean }>();
    for (const configuration of configurations) {
      for (const flow of relatedCloudFlows(flows.data ?? [], configuration)) {
        const current = seen.get(flow.id);
        seen.set(flow.id, { linked: Boolean(current?.linked) || flow.linked });
      }
    }
    return [...seen.values()];
  }, [configurations, flows.data]);
  const linkedFlows = generatedFlows.filter((flow) => flow.linked).length;
  const earlierVersions = generatedFlows.length - linkedFlows;
  const duplicateConfiguration = async (item: StudioConfiguration) => {
    try {
      const draft = draftFromConfiguration(item, runtime?.prompts ?? []);
      const id = await createConfiguration.mutateAsync(
        buildDraftConfigurationPayload({
          ...draft,
          name: `Copy of ${item.name}`,
          flowName: "",
        }),
      );
      notify("success", "Configuration duplicated", `"Copy of ${item.name}" was created as a draft.`);
      navigate(`/configurations/${id}`);
    } catch (error) {
      notify("error", "Could not duplicate the configuration", errorMessage(error, ""));
    }
  };
  const confirmPendingAction = async () => {
    if (!pendingAction) return;
    const { kind, item } = pendingAction;
    try {
      if (kind === "deactivate") {
        await deactivateConfiguration.mutateAsync(item.id);
        notify("success", "Configuration deactivated", `"${item.name}" is now inactive and hidden from the catalog.`);
      } else {
        await deleteConfiguration.mutateAsync(item.id);
        notify("success", "Configuration deleted", `"${item.name}" was removed from Dataverse.`);
      }
    } catch (error) {
      notify("error", kind === "deactivate" ? "Could not deactivate" : "Could not delete", errorMessage(error, ""));
    } finally {
      setPendingAction(null);
    }
  };
  const query = (new URLSearchParams(location.search).get("q") ?? "")
    .trim()
    .toLocaleLowerCase();
  const visibleConfigurations = query
    ? configurations.filter((item) =>
        [
          item.name,
          item.promptName,
          item.entity,
          item.outputEntity,
          item.outputField,
          item.flowName,
        ].some((value) => value?.toLocaleLowerCase().includes(query)),
      )
    : configurations;
  const published = configurations.filter(
    (item) => item.status === 100000001,
  ).length;
  const materializedFlows = configurations.filter((item) => item.flowId).length;
  const runsByConfiguration = new Map<string, number>();
  const latestRunByConfiguration = new Map<string, string>();
  for (const run of runtime?.runs ?? []) {
    runsByConfiguration.set(
      run.configurationId,
      (runsByConfiguration.get(run.configurationId) ?? 0) + 1,
    );
    if (run.timestamp) {
      const current = latestRunByConfiguration.get(run.configurationId);
      if (!current || run.timestamp > current) {
        latestRunByConfiguration.set(run.configurationId, run.timestamp);
      }
    }
  }
  return (
    <>
      <div className="studio-page catalog-page">
        <PageHeader
          eyebrow="Summary Studio"
          title="Summary configurations"
          description="Create and manage the Dataverse recipes that drive each generated summary."
        />

        <Toolbar aria-label="Configuration commands" className="catalog-command-bar">
          <ToolbarButton
            className="primary-command"
            appearance="primary"
            icon={<Sparkles />}
            onClick={() => navigate("/configurations/new")}
          >
            New configuration
          </ToolbarButton>
          <ToolbarButton
            appearance="subtle"
            icon={<RefreshCw />}
            onClick={() => refetch()}
            disabled={isFetching}
          >
            {isFetching ? "Refreshing" : "Refresh"}
          </ToolbarButton>
          <ToolbarDivider />
          <span className="catalog-command-hint">
            Stored in <code>csp_aisummaryconfig</code>
          </span>
        </Toolbar>

        <section className="catalog-status" aria-label="Studio summary">
          <span><b>{published}</b> published configurations</span>
          <span><b>{runtime?.cacheCount ?? 0}</b> available summaries</span>
          <span><b>{runtime?.runs.length ?? 0}</b> recorded runs</span>
          <span><b>{runtime?.accounts.length ?? 0}</b> demo records</span>
        </section>

        <section className="catalog-grid" aria-labelledby="catalog-grid-title">
          <div className="catalog-grid-heading">
            <div>
              <h2 id="catalog-grid-title">Configurations</h2>
              <p>Select a row to review or change its recipe.</p>
            </div>
            <Badge variant="outline">{visibleConfigurations.length} items</Badge>
          </div>
          <Table aria-label="Summary configurations" className="config-table" size="medium">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Configuration</TableHeaderCell>
                <TableHeaderCell>Source & destination</TableHeaderCell>
                <TableHeaderCell>Trigger</TableHeaderCell>
                <TableHeaderCell>Activity</TableHeaderCell>
                <TableHeaderCell>Status</TableHeaderCell>
                <TableHeaderCell aria-label="Open" />
              </TableRow>
            </TableHeader>
            <TableBody>
              {visibleConfigurations.map((item) => (
                <TableRow key={item.id || item.name}>
                  <TableCell>
                    <TableCellLayout media={<span className="entity-icon"><Database /></span>}>
                      <Link className="configuration-link" to={`/configurations/${item.id}`}>
                        <b>{item.name}</b>
                        <small>{item.promptName || "AI Prompt"}</small>
                      </Link>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <span className="table-value">
                      <b>{item.entity === "account" ? "Account" : item.entity}</b>
                      <small>{item.outputEntity}.{item.outputField}</small>
                    </span>
                  </TableCell>
                  <TableCell>
                    <span className="table-value">
                      <b>{item.mode === 100000000 ? "Per record" : "Consolidated"}</b>
                      <small>
                        {item.flowId
                          ? item.flowDisplayName || item.flowName || "Flow generated"
                          : item.flowName
                            ? `${item.flowName} · pending`
                            : "Flow pending"}
                      </small>
                    </span>
                  </TableCell>
                  <TableCell>
                    <span className="table-value">
                      <b>{runsByConfiguration.get(item.id) ?? 0} runs</b>
                      <small>
                        {(() => {
                          const lastActivity =
                            item.lastRun ?? latestRunByConfiguration.get(item.id);
                          return lastActivity
                            ? `Last run ${formatRelativeDate(lastActivity)}`
                            : "Not run yet";
                        })()}
                      </small>
                    </span>
                  </TableCell>
                  <TableCell>
                    <Badge
                      variant="outline"
                      className={item.status === 100000001 ? "status-live" : "status-draft"}
                    >
                      {item.status === 100000001 ? "Published" : "Draft"}
                    </Badge>
                  </TableCell>
                  <TableCell>
                    <span className="row-actions">
                      <Menu positioning="below-end">
                        <MenuTrigger disableButtonEnhancement>
                          <Button
                            variant="ghost"
                            size="icon-sm"
                            aria-label={`Actions for ${item.name}`}
                          >
                            <More />
                          </Button>
                        </MenuTrigger>
                        <MenuPopover>
                          <MenuList>
                            <MenuItem
                              icon={<Braces />}
                              onClick={() => navigate(`/configurations/${item.id}`)}
                            >
                              Open
                            </MenuItem>
                            <MenuItem
                              icon={<Copy />}
                              disabled={createConfiguration.isPending}
                              onClick={() => duplicateConfiguration(item)}
                            >
                              Duplicate
                            </MenuItem>
                            <MenuItem
                              icon={<Dismiss />}
                              onClick={() => setPendingAction({ kind: "deactivate", item })}
                            >
                              Deactivate
                            </MenuItem>
                            <MenuItem
                              icon={<Delete />}
                              onClick={() => setPendingAction({ kind: "delete", item })}
                            >
                              Delete
                            </MenuItem>
                          </MenuList>
                        </MenuPopover>
                      </Menu>
                      <Button
                        variant="ghost"
                        size="icon-sm"
                        aria-label={`Open ${item.name}`}
                        onClick={() => navigate(`/configurations/${item.id}`)}
                      >
                        <ChevronRight />
                      </Button>
                    </span>
                  </TableCell>
                </TableRow>
              ))}
              {visibleConfigurations.length === 0 && (
                <TableRow>
                  <TableCell colSpan={6}>
                    <div className="catalog-empty-state">
                      <Search />
                      <span>
                        <b>No configurations match your search.</b>
                        <small>Try a name, table, prompt, or flow.</small>
                      </span>
                    </div>
                  </TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>
        </section>

        <section className="platform-strip">
          <span className="platform-strip-icon"><Flow /></span>
          <div>
            <b>Power Platform pipeline</b>
            <p>Dataverse configuration → AI Prompt → generated cloud flow</p>
          </div>
          <span>
            {configurations.length === 0
              ? "No recipes yet"
              : flows.isLoading
                ? "Checking generated flows…"
                : generatedFlows.length === 0
                  ? materializedFlows === 0
                    ? "Flow generation pending in the backend"
                    : `${materializedFlows} of ${configurations.length} flows linked`
                  : `${generatedFlows.length} flow${generatedFlows.length === 1 ? "" : "s"} generated · ${linkedFlows} linked${
                      earlierVersions > 0
                        ? ` · ${earlierVersions} earlier version${earlierVersions === 1 ? "" : "s"} still active`
                        : ""
                    }`}
          </span>
        </section>
      </div>
      <Dialog
        open={Boolean(pendingAction)}
        onOpenChange={(_, dialogData) => {
          if (!dialogData.open) setPendingAction(null);
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {pendingAction?.kind === "delete"
                ? "Delete configuration?"
                : "Deactivate configuration?"}
            </DialogTitle>
            <DialogContent>
              {pendingAction?.kind === "delete" ? (
                <>
                  <b>{pendingAction.item.name}</b> will be removed from Dataverse.
                  Generated flows and recorded runs are not deleted; retire the flow
                  in Power Automate if it should stop running.
                </>
              ) : (
                <>
                  <b>{pendingAction?.item.name}</b> will be set inactive and disappear
                  from the catalog. The generated flow keeps running until it is
                  turned off in Power Automate.
                </>
              )}
            </DialogContent>
            <DialogActions>
              <Button variant="outline" onClick={() => setPendingAction(null)}>
                Cancel
              </Button>
              <Button
                variant={pendingAction?.kind === "delete" ? "destructive" : "default"}
                disabled={deactivateConfiguration.isPending || deleteConfiguration.isPending}
                onClick={confirmPendingAction}
              >
                {pendingAction?.kind === "delete" ? "Delete" : "Deactivate"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}

const steps = [
  {
    title: "Data and context",
    stage: "Context",
    subtitle: "What it should read",
    icon: Database,
  },
  {
    title: "Prompt and model",
    stage: "Prompt",
    subtitle: "How it should reason",
    icon: Bot,
  },
  {
    title: "Destination",
    stage: "Destination",
    subtitle: "Where it should write",
    icon: Target,
  },
  {
    title: "Execution",
    stage: "Automation",
    subtitle: "When it should run",
    icon: Flow,
  },
  {
    title: "Review and publish",
    stage: "Publish",
    subtitle: "What will be generated",
    icon: CheckCircle,
  },
];

function LabeledField({
  label,
  children,
  hint,
}: {
  label: string;
  children: ReactNode;
  hint?: string;
}) {
  return (
    <label className="studio-field">
      <span>{label}</span>
      {children}
      {hint && <small>{hint}</small>}
    </label>
  );
}

type PaneKind =
  | "table"
  | "execution"
  | "prompt"
  | "model"
  | "fields"
  | "records"
  | "relationships"
  | "contextQuery"
  | "outputTable"
  | "outputField";

type BuilderDraft = SummaryConfigurationDraft;

function draftFromConfiguration(
  config?: StudioConfiguration,
  prompts: StudioPrompt[] = [],
): BuilderDraft {
  const prompt =
    prompts.find((item) => item.id === config?.promptId) ?? prompts[0];
  const entity = config?.entity ?? "account";
  const sourceFields = config?.sourceFields.length
    ? config.sourceFields
    : ["name"];
  const relationships = config?.relationships ?? [];
  const maxRecords = config?.maxRecords ?? 5000;
  return {
    name: config?.name ?? "New summary",
    entity,
    entitySetName:
      config?.entitySetName ?? getTargetEntityIdentity(entity).entitySet,
    entityIdField:
      config?.entityIdField ?? getTargetEntityIdentity(entity).idField,
    sourceFields,
    outputEntity: config?.outputEntity ?? "account",
    outputEntitySetName:
      config?.outputEntitySetName ||
      getTargetEntityIdentity(config?.outputEntity ?? "account").entitySet,
    outputField: config?.outputField ?? "csp_aisummary",
    model: config?.model ?? prompt?.model ?? "gpt-4.1-mini",
    mode: config?.mode ?? 100000000,
    promptId: prompt?.id ?? "",
    promptName: prompt?.name ?? "Select AI Prompt",
    promptKey: prompt?.key ?? "",
    promptContent: prompt?.content ?? "",
    flowName: config?.flowName ?? "",
    relationships,
    inputMappings: config?.inputMappings ?? {},
    triggerColumns: config?.triggerColumns.length
      ? config.triggerColumns
      : ["name", "revenue", "description", "primarycontactid"],
    triggerType: "dataverse.update",
    preserveHistory: false,
    saveMetadata: true,
    queryMode: config?.queryMode ?? 100000002,
    systemViewId: config?.systemViewId,
    userViewId: config?.userViewId,
    maxRecords,
    fetchXml:
      config?.fetchXml ||
      buildRecordSelectionFetchXml(entity, maxRecords, config?.entityIdField),
    relatedFetchXml:
      config?.relatedFetchXml ||
      buildContextFetchXml(
        entity,
        sourceFields,
        relationships,
        config?.entityIdField,
      ),
  };
}

function DataStep({
  onEdit,
  draft,
  accounts,
}: {
  onEdit: (pane: PaneKind) => void;
  draft: BuilderDraft;
  accounts: StudioAccount[];
}) {
  const tables = useDataverseTables();
  const entityLabel =
    tables.data?.find((table) => table.logicalName === draft.entity)
      ?.displayName ?? draft.entity;
  const relatedEntity = String(draft.relationships[0]?.entity ?? "");
  const relatedLabel =
    tables.data?.find((table) => table.logicalName === relatedEntity)
      ?.displayName ?? relatedEntity;
  const relationshipCount = draft.relationships.length;
  const recordSelectionLabel =
    draft.queryMode === 100000000
      ? "System view"
      : draft.queryMode === 100000001
        ? "Personal view"
        : "Filters or FetchXML";
  const relatedMode = Number(
    draft.relationships[0]?.selectionMode ?? 100000002,
  );
  const relatedSelectionLabel =
    relatedMode === 100000000
      ? "System view"
      : relatedMode === 100000001
        ? "Personal view"
        : "Filters or FetchXML";
  return (
    <div className="editor-section data-context-step">
      <SectionIntro
        number="01"
        title="Data and context"
        text="Configure the record set first, then define the Dataverse context sent to the prompt for each record."
      />
      <div className="configuration-sections">
        <section
          className="configuration-section"
          aria-label="1. Records to summarize"
        >
          <header>
            <span>1</span>
            <div>
              <h3>Records to summarize</h3>
              <p>Define which Dataverse rows enter the generated flow.</p>
            </div>
          </header>
          <div className="configuration-list">
            <article>
              <span className="configuration-icon">
                <Database />
              </span>
              <div className="configuration-label">
                <b>Source table</b>
                <code>csp_targetentity</code>
              </div>
              <div className="configuration-value">
                <b>{entityLabel}</b>
                <code>{draft.entity}</code>
              </div>
              <Button
                type="button"
                variant="outline"
                size="sm"
                aria-label="Change source table"
                onClick={() => onEdit("table")}
              >
                Change
              </Button>
            </article>
            <article>
              <span className="configuration-icon">
                <Search />
              </span>
              <div className="configuration-label">
                <b>Record selection</b>
                <code>csp_fetchxml</code>
              </div>
              <div className="configuration-value">
                <b>{recordSelectionLabel}</b>
                <small>{accounts.length} demo records</small>
                <small>
                  Maximum {draft.maxRecords.toLocaleString("en-US")}
                </small>
              </div>
              <Button
                type="button"
                variant="outline"
                size="sm"
                aria-label="Configure record FetchXML"
                onClick={() => onEdit("records")}
              >
                Configure
              </Button>
            </article>
            <article>
              <span className="configuration-icon">
                <Blocks />
              </span>
              <div className="configuration-label">
                <b>Summary mode</b>
                <code>csp_mode</code>
              </div>
              <div className="configuration-value">
                <b>
                  {draft.mode === 100000000
                    ? "One summary per record"
                    : "Consolidated summary"}
                </b>
                <code>
                  {draft.mode === 100000000 ? "PerRecord" : "Aggregate"}
                </code>
              </div>
              <Button
                type="button"
                variant="outline"
                size="sm"
                aria-label="Change summary mode"
                onClick={() => onEdit("execution")}
              >
                Change
              </Button>
            </article>
          </div>
        </section>

        <section
          className="configuration-section"
          aria-label="2. Context for each record"
        >
          <header>
            <span>2</span>
            <div>
              <h3>Context for each record</h3>
              <p>Control exactly what is passed to the selected AI Prompt.</p>
            </div>
          </header>
          <div className="configuration-list">
            <article>
              <span className="configuration-icon">
                <Database />
              </span>
              <div className="configuration-label">
                <b>Source fields</b>
                <code>csp_sourcefields</code>
              </div>
              <div className="configuration-value configuration-tokens">
                {draft.sourceFields.map((field) => (
                  <code key={field}>{field}</code>
                ))}
              </div>
              <Button
                type="button"
                variant="outline"
                size="sm"
                aria-label="Configure source fields"
                onClick={() => onEdit("fields")}
              >
                Select fields
              </Button>
            </article>
            <article>
              <span className="configuration-icon">
                <GitBranch />
              </span>
              <div className="configuration-label">
                <b>Related records</b>
                <code>csp_relationships</code>
              </div>
              <div className="configuration-value">
                <b>
                  {relationshipCount
                    ? relatedSelectionLabel
                    : "No related records"}
                </b>
                <code>
                  {relationshipCount
                    ? `${relatedLabel} · ${relatedEntity}`
                    : "none"}
                </code>
              </div>
              <Button
                type="button"
                variant="outline"
                size="sm"
                aria-label="Configure related records"
                onClick={() => onEdit("relationships")}
              >
                Configure
              </Button>
            </article>
            <article>
              <span className="configuration-icon">
                <Code />
              </span>
              <div className="configuration-label">
                <b>Context query</b>
                <code>csp_relatedfetchxml</code>
              </div>
              <div className="configuration-value">
                <b>FetchXML per record</b>
                <small>
                  {draft.sourceFields.length} fields · {relationshipCount}{" "}
                  related table
                </small>
              </div>
              <Button
                type="button"
                variant="outline"
                size="sm"
                aria-label="Configure context FetchXML"
                onClick={() => onEdit("contextQuery")}
              >
                Edit FetchXML
              </Button>
            </article>
          </div>
        </section>
      </div>
      <div className="configuration-result">
        <CheckCircle />
        <span>
          <b>Compiled context</b>
          <small>
            For every matching <code>{draft.entity}</code>, the flow runs{" "}
            <code>csp_relatedfetchxml</code> and passes its result to the
            prompt.
          </small>
        </span>
        <Button
          type="button"
          variant="outline"
          size="sm"
          onClick={() => onEdit("contextQuery")}
        >
          Review query
        </Button>
      </div>
    </div>
  );
}

const defaultPromptWizardAnswers: PromptWizardAnswers = {
  objective:
    "Give account managers a concise operational briefing before a customer conversation.",
  audience: "Account managers",
  format: "Structured sections",
  length: "Concise",
  tone: "Direct and professional",
  language: "English",
  includeActions: true,
  citeEvidence: true,
  handleMissingData: true,
};

function PromptAssistantPane({
  draft,
  onClose,
  onApply,
}: {
  draft: BuilderDraft;
  onClose: () => void;
  onApply: (prompt: string) => void;
}) {
  const [step, setStep] = useState(1);
  const [answers, setAnswers] = useState(defaultPromptWizardAnswers);
  const [generatedPrompt, setGeneratedPrompt] = useState("");
  const update = <K extends keyof PromptWizardAnswers>(
    key: K,
    value: PromptWizardAnswers[K],
  ) => setAnswers((current) => ({ ...current, [key]: value }));
  const generate = () => {
    setGeneratedPrompt(
      buildPromptDraft({
        answers,
        contextVariable: "{{account_context}}",
        entityLabel: draft.entity === "account" ? "Account" : draft.entity,
        fields: draft.sourceFields,
        relatedSources: draft.relationships
          .map((relationship) => String(relationship.entity ?? ""))
          .filter(Boolean),
      }),
    );
    setStep(3);
  };

  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close prompt design assistant"
        onClick={onClose}
      />
      <aside
        className="edit-pane prompt-wizard-pane"
        aria-label="Prompt design assistant"
      >
        <header>
          <div>
            <p className="micro-label text-brand-blue">
              AI-assisted prompt builder
            </p>
            <h2>Design a better prompt</h2>
          </div>
          <Button
            type="button"
            variant="ghost"
            size="icon"
            aria-label="Close prompt design assistant"
            onClick={onClose}
          >
            <Dismiss />
          </Button>
        </header>
        <nav
          className="prompt-wizard-progress"
          aria-label="Prompt wizard progress"
        >
          {["Purpose", "Output", "Review"].map((label, index) => (
            <span
              key={label}
              className={
                step === index + 1
                  ? "active"
                  : step > index + 1
                    ? "complete"
                    : ""
              }
            >
              <i>{step > index + 1 ? <Check /> : index + 1}</i>
              {label}
            </span>
          ))}
        </nav>
        <section className="prompt-wizard-body">
          {step === 1 && (
            <div className="prompt-wizard-stage">
              <p className="micro-label text-brand-blue">Step 1 · Purpose</p>
              <h3>What should this summary achieve?</h3>
              <p className="wizard-lead">
                Describe the business outcome. The assistant already knows which
                Dataverse fields and related records are available.
              </p>
              <LabeledField label="Summary objective">
                <Textarea
                  resize="vertical"
                  aria-label="Summary objective"
                  value={answers.objective}
                  onChange={(event) => update("objective", event.target.value)}
                  rows={4}
                />
              </LabeledField>
              <div className="grid grid-cols-2 gap-3">
                <LabeledField label="Audience">
                  <Select
                    aria-label="Audience"
                    value={answers.audience}
                    onChange={(event) => update("audience", event.target.value)}
                  >
                    <option>Account managers</option>
                    <option>Executives</option>
                    <option>Service agents</option>
                    <option>Operations teams</option>
                  </Select>
                </LabeledField>
                <LabeledField label="Language">
                  <Select
                    aria-label="Output language"
                    value={answers.language}
                    onChange={(event) => update("language", event.target.value)}
                  >
                    <option>English</option>
                    <option>Spanish</option>
                    <option>Use the record language</option>
                  </Select>
                </LabeledField>
              </div>
              <div className="wizard-context">
                <Database />
                <div>
                  <b>Context is already connected</b>
                  <p>
                    {draft.sourceFields.length} fields and{" "}
                    {draft.relationships.length} related source
                    {draft.relationships.length === 1 ? "" : "s"} will be
                    available as <code>{"{{account_context}}"}</code>.
                  </p>
                </div>
              </div>
            </div>
          )}
          {step === 2 && (
            <div className="prompt-wizard-stage">
              <p className="micro-label text-brand-blue">Step 2 · Output</p>
              <h3>Shape the answer</h3>
              <p className="wizard-lead">
                Choose a dependable output pattern and the safeguards the
                generated prompt must enforce.
              </p>
              <div className="grid grid-cols-2 gap-3">
                <LabeledField label="Format">
                  <Select
                    aria-label="Output format"
                    value={answers.format}
                    onChange={(event) => update("format", event.target.value)}
                  >
                    <option>Structured sections</option>
                    <option>Bullet points</option>
                    <option>Short narrative</option>
                  </Select>
                </LabeledField>
                <LabeledField label="Length">
                  <Select
                    aria-label="Summary length"
                    value={answers.length}
                    onChange={(event) => update("length", event.target.value)}
                  >
                    <option>Concise</option>
                    <option>Standard</option>
                    <option>Detailed</option>
                  </Select>
                </LabeledField>
              </div>
              <LabeledField label="Tone">
                <Select
                  aria-label="Tone"
                  value={answers.tone}
                  onChange={(event) => update("tone", event.target.value)}
                >
                  <option>Direct and professional</option>
                  <option>Executive and strategic</option>
                  <option>Clear and supportive</option>
                  <option>Neutral and factual</option>
                </Select>
              </LabeledField>
              <div className="wizard-guardrails">
                <p className="field-caption">Guardrails</p>
                <label>
                  <Checkbox
                    checked={answers.includeActions}
                    onChange={(_, data) =>
                      update("includeActions", data.checked === true)
                    }
                  />
                  <span>
                    <b>Recommend next actions</b>
                    <small>
                      Turn the summary into an operational briefing.
                    </small>
                  </span>
                </label>
                <label>
                  <Checkbox
                    checked={answers.citeEvidence}
                    onChange={(_, data) =>
                      update("citeEvidence", data.checked === true)
                    }
                  />
                  <span>
                    <b>Ground important statements</b>
                    <small>
                      Connect claims to values and events in the context.
                    </small>
                  </span>
                </label>
                <label>
                  <Checkbox
                    checked={answers.handleMissingData}
                    onChange={(_, data) =>
                      update("handleMissingData", data.checked === true)
                    }
                  />
                  <span>
                    <b>Handle missing data explicitly</b>
                    <small>Never fill gaps with invented information.</small>
                  </span>
                </label>
              </div>
            </div>
          )}
          {step === 3 && (
            <div className="prompt-wizard-stage">
              <p className="micro-label text-brand-blue">Step 3 · Review</p>
              <h3>Review the generated prompt</h3>
              <p className="wizard-lead">
                The proposal combines your choices with the active Dataverse
                context. Edit anything before applying it.
              </p>
              <Textarea
                resize="vertical"
                className="generated-prompt"
                aria-label="Generated prompt draft"
                value={generatedPrompt}
                onChange={(event) => setGeneratedPrompt(event.target.value)}
              />
              <div className="wizard-validation">
                <CheckCircle />
                <span>
                  <b>Ready to use</b>
                  <small>
                    Context variable preserved · evidence rules included ·
                    output structure defined
                  </small>
                </span>
              </div>
            </div>
          )}
        </section>
        <footer>
          {step > 1 ? (
            <Button
              type="button"
              variant="outline"
              onClick={() => setStep((current) => current - 1)}
            >
              Back
            </Button>
          ) : (
            <Button type="button" variant="outline" onClick={onClose}>
              Cancel
            </Button>
          )}
          <span className="flex-1" />
          {step === 1 ? (
            <Button
              type="button"
              onClick={() => setStep(2)}
              disabled={!answers.objective.trim()}
            >
              Continue to output <ArrowRight data-icon="inline-end" />
            </Button>
          ) : step === 2 ? (
            <Button type="button" onClick={generate}>
              <Sparkles data-icon="inline-start" />
              Generate prompt draft
            </Button>
          ) : (
            <Button
              type="button"
              onClick={() => onApply(generatedPrompt)}
              disabled={!generatedPrompt.trim()}
            >
              <Check data-icon="inline-start" />
              Use this prompt
            </Button>
          )}
        </footer>
      </aside>
    </>
  );
}

function PromptStep({
  onEdit,
  draft,
  onSavePrompt,
  savingPrompt,
}: {
  onEdit: (pane: PaneKind) => void;
  draft: BuilderDraft;
  onSavePrompt: (content: string) => Promise<void>;
  savingPrompt: boolean;
}) {
  const [assistantOpen, setAssistantOpen] = useState(false);
  const [promptContent, setPromptContent] = useState(draft.promptContent);
  const [saveState, setSaveState] = useState<"idle" | "saved" | "error">("idle");
  const [saveError, setSaveError] = useState("");
  const savePrompt = async () => {
    try {
      await onSavePrompt(promptContent);
      setSaveState("saved");
      setSaveError("");
      window.setTimeout(() => setSaveState("idle"), 3000);
    } catch (error) {
      setSaveState("error");
      setSaveError(error instanceof Error ? error.message : "The prompt could not be saved.");
    }
  };
  const dirty = promptContent !== draft.promptContent;
  const estimatedTokens = estimateRecipeTokens({
    promptContent,
    sourceFields: draft.sourceFields,
    relationships: draft.relationships,
  });

  return (
    <div className="editor-section">
      <SectionIntro
        number="02"
        title="Prompt and model"
        text="Keep the maker experience simple without hiding the architecture and governance decisions that matter."
      />
      <div className="grid gap-4 md:grid-cols-[1fr_220px]">
        <LabeledField label="AI Prompt">
          <button
            type="button"
            className="field-control w-full text-left"
            aria-label="Edit AI Prompt"
            onClick={() => onEdit("prompt")}
          >
            <Bot />
            <span>
              <b>{draft.promptName}</b>
              <code>{draft.promptKey}</code>
            </span>
            <ChevronRight />
          </button>
        </LabeledField>
        <LabeledField label="Model">
          <button
            type="button"
            className="field-control compact w-full text-left"
            aria-label="Edit model"
            onClick={() => onEdit("model")}
          >
            <Sparkles />
            <span>
              <b>
                {draft.model.replace("gpt-", "GPT-").replace("-mini", " mini")}
              </b>
              <code>AI Builder model</code>
            </span>
          </button>
        </LabeledField>
      </div>
      <div className="studio-field">
        <span>Prompt instructions</span>
        <div
          className="prompt-instructions-editor"
          role="group"
          aria-label="Prompt instructions editor"
        >
          <div className="prompt-editor-assistant">
            <span>
              <Sparkles />
            </span>
            <p>
              <b>Prompt assistant</b>
              <small>
                Build from purpose, audience, format, and the active Dataverse
                context.
              </small>
            </p>
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => setAssistantOpen(true)}
            >
              <Sparkles data-icon="inline-start" />
              Design with assistant
            </Button>
          </div>
          <Textarea
            resize="none"
            className="prompt-fluent-textarea w-full font-mono text-sm leading-6"
            value={promptContent}
            onChange={(event) => setPromptContent(event.target.value)}
            aria-label="Prompt instructions"
          />
          <div className="prompt-editor-footer">
            <span className="prompt-runtime-label">Runtime inputs</span>
            <code>{"{{account_context}}"}</code>
            <code>{"{{generated_at}}"}</code>
            <span className="prompt-cost">
              <span>
                Estimate · ~{estimatedTokens.toLocaleString("en-US")} tokens per
                run
              </span>
            </span>
            <Button
              type="button"
              size="sm"
              variant={dirty ? "default" : "outline"}
              disabled={savingPrompt || !draft.promptId || !dirty}
              onClick={savePrompt}
            >
              {savingPrompt
                ? "Saving prompt…"
                : saveState === "saved" && !dirty
                  ? "Saved"
                  : "Save prompt"}
            </Button>
          </div>
        </div>
        <p className="prompt-save-hint" aria-live="polite">
          {saveState === "error" ? (
            <span className="prompt-save-error">{saveError}</span>
          ) : dirty ? (
            <>
              Unsaved changes. <b>Save prompt</b> updates the shared AI Prompt{" "}
              <code>{draft.promptKey || draft.promptName}</code> in Dataverse, so
              every configuration bound to it uses the new instructions.
            </>
          ) : (
            <>
              These instructions are stored on the AI Prompt record{" "}
              <code>{draft.promptKey || draft.promptName}</code>. Edit them here or{" "}
              <Link to="/prompts">manage all prompts</Link>.
            </>
          )}
        </p>
      </div>
      {assistantOpen && (
        <PromptAssistantPane
          draft={draft}
          onClose={() => setAssistantOpen(false)}
          onApply={(prompt) => {
            setPromptContent(prompt);
            setAssistantOpen(false);
          }}
        />
      )}
    </div>
  );
}

function DestinationStep({
  draft,
  onEdit,
  onChange,
}: {
  draft: BuilderDraft;
  onEdit: (pane: PaneKind) => void;
  onChange: (changes: Partial<BuilderDraft>) => void;
}) {
  const tables = useDataverseTables();
  const columns = useDataverseColumns(
    draft.outputEntity,
    draft.outputEntitySetName ||
      getTargetEntityIdentity(draft.outputEntity).entitySet,
  );
  const sourceLabel =
    tables.data?.find((table) => table.logicalName === draft.entity)
      ?.displayName ?? draft.entity;
  const outputLabel =
    tables.data?.find((table) => table.logicalName === draft.outputEntity)
      ?.displayName ?? draft.outputEntity;
  const outputFieldLabel =
    columns.data?.find((column) => column.logicalName === draft.outputField)
      ?.displayName ?? draft.outputField;
  return (
    <div className="editor-section">
      <SectionIntro
        number="03"
        title="Destination"
        text="Write the summary back to Dataverse so it can be reused in views, forms, agents, and other automations."
      />
      <div className="destination-map">
        <div>
          <Database />
          <span>
            <small>Source record</small>
            <b>{sourceLabel}</b>
            <code>
              {draft.entity}.{draft.entityIdField}
            </code>
          </span>
        </div>
        <ArrowRight />
        <div className="active">
          <Sparkles />
          <span>
            <small>Generated result</small>
            <b>{outputFieldLabel}</b>
            <code>
              {draft.outputEntity}.{draft.outputField}
            </code>
          </span>
        </div>
      </div>
      <div className="grid gap-4 md:grid-cols-2">
        <LabeledField label="Destination table">
          <button
            type="button"
            className="field-control w-full text-left"
            aria-label="Edit destination table"
            onClick={() => onEdit("outputTable")}
          >
            <Database />
            <span>
              <b>{outputLabel}</b>
              <code>{draft.outputEntity}</code>
            </span>
            <ChevronRight />
          </button>
        </LabeledField>
        <LabeledField label="Destination column">
          <button
            type="button"
            className="field-control w-full text-left"
            aria-label="Edit destination column"
            onClick={() => onEdit("outputField")}
          >
            <Target />
            <span>
              <b>{outputFieldLabel}</b>
              <code>{draft.outputField}</code>
            </span>
            <ChevronRight />
          </button>
        </LabeledField>
      </div>
      <div className="option-list">
        <button
          type="button"
          aria-pressed={!draft.preserveHistory}
          className={!draft.preserveHistory ? "selected" : ""}
          onClick={() => onChange({ preserveHistory: false })}
        >
          <span className="radio-dot" />
          <div>
            <b>Overwrite the current summary</b>
            <small>Keep one current result on the source record.</small>
          </div>
          {!draft.preserveHistory && <Check />}
        </button>
        <button
          type="button"
          aria-pressed={draft.preserveHistory}
          className={draft.preserveHistory ? "selected" : ""}
          onClick={() => onChange({ preserveHistory: true })}
        >
          <span className="radio-dot" />
          <div>
            <b>Preserve history</b>
            <small>Create a cache entry for every run.</small>
          </div>
          {draft.preserveHistory && <Check />}
        </button>
      </div>
      <div className="inline-setting">
        <div>
          <b>Store generation metadata</b>
          <small>Model, tokens, duration, and configuration version.</small>
        </div>
        <Switch
          checked={draft.saveMetadata}
          onCheckedChange={(checked) => onChange({ saveMetadata: checked })}
          aria-label="Store generation metadata"
        />
      </div>
    </div>
  );
}

function triggerPresentation(type: string, columns: string[]) {
  if (type === "schedule.daily")
    return {
      label: "Daily schedule",
      detail: "Every day · Europe/Madrid",
      pattern: "Scheduled",
    };
  if (type === "manual")
    return {
      label: "On-demand execution",
      detail: "Explicit invocation",
      pattern: "On demand",
    };
  return {
    label: "When an account changes",
    detail: columns.slice(0, 3).join(", "),
    pattern: "Dataverse event",
  };
}

function FlowDiagram({ draft }: { draft: BuilderDraft }) {
  const trigger = triggerPresentation(draft.triggerType, draft.triggerColumns);
  return (
    <div className="generated-flow">
      <div className="flow-node trigger">
        <span>
          <Clock />
        </span>
        <div>
          <small>Trigger</small>
          <b>{trigger.label}</b>
          <code>{trigger.detail}</code>
        </div>
      </div>
      <i />
      <div className="flow-node">
        <span>
          <Database />
        </span>
        <div>
          <small>Dataverse</small>
          <b>Query Dataverse</b>
          <code>FetchXML + relationships</code>
        </div>
      </div>
      <i />
      <div className="flow-node">
        <span>
          <Bot />
        </span>
        <div>
          <small>AI Builder</small>
          <b>Run AI Prompt</b>
          <code>{draft.promptKey}</code>
        </div>
      </div>
      <i />
      <div className="flow-node">
        <span>
          <Target />
        </span>
        <div>
          <small>Dataverse</small>
          <b>Save summary</b>
          <code>
            {draft.outputEntity}.{draft.outputField}
          </code>
        </div>
      </div>
    </div>
  );
}

function ExecutionStep({
  draft,
  onChange,
}: {
  draft: BuilderDraft;
  onChange: (changes: Partial<BuilderDraft>) => void;
}) {
  const trigger = triggerPresentation(draft.triggerType, draft.triggerColumns);
  return (
    <div className="editor-section">
      <SectionIntro
        number="04"
        title="Execution"
        text="Select a governed pattern. Summary Studio generates the trigger and actions without turning the app into a universal flow designer."
      />
      <div className="trigger-options">
        <button
          type="button"
          aria-label="When an account changes"
          aria-pressed={draft.triggerType === "dataverse.update"}
          className={draft.triggerType === "dataverse.update" ? "selected" : ""}
          onClick={() => onChange({ triggerType: "dataverse.update" })}
        >
          <span>
            <GitBranch />
          </span>
          <div>
            <b>When an account changes</b>
            <small>Only when a field used by the context changes.</small>
          </div>
          {draft.triggerType === "dataverse.update" && <Check />}
        </button>
        <button
          type="button"
          aria-label="Scheduled"
          aria-pressed={draft.triggerType === "schedule.daily"}
          className={draft.triggerType === "schedule.daily" ? "selected" : ""}
          onClick={() => onChange({ triggerType: "schedule.daily" })}
        >
          <span>
            <Clock />
          </span>
          <div>
            <b>Scheduled</b>
            <small>Daily, weekly, or a custom recurrence.</small>
          </div>
          {draft.triggerType === "schedule.daily" && <Check />}
        </button>
        <button
          type="button"
          aria-label="On demand"
          aria-pressed={draft.triggerType === "manual"}
          className={draft.triggerType === "manual" ? "selected" : ""}
          onClick={() => onChange({ triggerType: "manual" })}
        >
          <span>
            <Play />
          </span>
          <div>
            <b>On demand</b>
            <small>From a command or as a child flow.</small>
          </div>
          {draft.triggerType === "manual" && <Check />}
        </button>
      </div>
      <div className="change-filter">
        <div>
          <p className="field-caption">Trigger columns</p>
          <div>
            {draft.triggerColumns.map((x) => (
              <span key={x}>{x}</span>
            ))}
          </div>
        </div>
        <aside>
          <RefreshCw />
          <span>
            <b>Loop prevention active</b>
            <small>The output column is excluded from the trigger.</small>
          </span>
        </aside>
      </div>
      <div className="flow-preview">
        <div className="flex items-center justify-between">
          <div>
            <p className="field-caption">Compiled definition</p>
            <h3>Flow contract for the backend</h3>
          </div>
          <Badge variant="outline">Pattern · {trigger.pattern}</Badge>
        </div>
        <FlowDiagram draft={draft} />
      </div>
    </div>
  );
}

function ReviewStep({
  draft,
  validationErrors,
}: {
  draft: BuilderDraft;
  validationErrors: string[];
}) {
  const trigger = triggerPresentation(draft.triggerType, draft.triggerColumns);
  const relatedEntities = draft.relationships
    .map((relationship) => String(relationship.entity ?? ""))
    .filter(Boolean);
  const relationshipCount = draft.relationships.length;
  const inputCount = Object.keys(draft.inputMappings).length;
  return (
    <div className="editor-section">
      <SectionIntro
        number="05"
        title="Review and publish"
        text="The recipe is complete. Publishing creates or updates only the assets required by the solution."
      />
      <div className="review-grid">
        <ReviewCard
          icon={<Database />}
          title="Data"
          value={
            relatedEntities.length
              ? `${draft.entity} + ${relatedEntities.join(", ")}`
              : draft.entity
          }
          detail={`${draft.sourceFields.length} field${draft.sourceFields.length === 1 ? "" : "s"} · ${relationshipCount} relationship${relationshipCount === 1 ? "" : "s"}`}
        />
        <ReviewCard
          icon={<Bot />}
          title="Generation"
          value={draft.promptName}
          detail={`${draft.model} · ${inputCount} input${inputCount === 1 ? "" : "s"}`}
        />
        <ReviewCard
          icon={<Target />}
          title="Destination"
          value={`${draft.outputEntity}.${draft.outputField}`}
          detail={`${draft.preserveHistory ? "Preserve history" : "Overwrite"} · metadata ${draft.saveMetadata ? "enabled" : "disabled"}`}
        />
        <ReviewCard
          icon={<Flow />}
          title="Execution"
          value={trigger.label}
          detail={
            draft.triggerType === "dataverse.update"
              ? "Loop prevention prepared"
              : trigger.detail
          }
        />
      </div>
      <div className="publish-contract">
        <div>
          <p className="field-caption">Prepared assets</p>
          <div className="asset-list">
            <span>
              <CheckCircle />1 configuration record{" "}
              <code>csp_aisummaryconfig</code>
            </span>
            <span>
              <CheckCircle />1 versioned definition for the flow generator{" "}
              <code>{draft.flowName || "Generated after publishing"}</code>
            </span>
            <span>
              <LinkIcon />1 binding to an existing AI Prompt{" "}
              <code>{draft.promptKey || "Not selected"}</code>
            </span>
          </div>
        </div>
        <aside aria-live="polite">
          {validationErrors.length === 0 ? (
            <>
              <p>
                <CheckCircle /> Validation complete
              </p>
              <small>
                Valid schema; flow materialization is handled by the backend.
              </small>
            </>
          ) : (
            <>
              <p>
                Resolve {validationErrors.length} item
                {validationErrors.length === 1 ? "" : "s"}
              </p>
              <ul className="mt-2 list-disc space-y-1 pl-4 text-xs">
                {validationErrors.map((error) => (
                  <li key={error}>{error}</li>
                ))}
              </ul>
            </>
          )}
        </aside>
      </div>
    </div>
  );
}

function ReviewCard({
  icon,
  title,
  value,
  detail,
}: {
  icon: ReactNode;
  title: string;
  value: string;
  detail: string;
}) {
  return (
    <div className="review-card">
      <span>{icon}</span>
      <p>{title}</p>
      <b>{value}</b>
      <small>{detail}</small>
    </div>
  );
}
function SectionIntro({
  number,
  title,
  text,
}: {
  number: string;
  title: string;
  text: string;
}) {
  return (
    <div className="section-intro">
      <span>{number}</span>
      <div>
        <h2>{title}</h2>
        <p>{text}</p>
      </div>
    </div>
  );
}

function isGuidLike(value: unknown) {
  return (
    typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)
  );
}

function formatPreviewValue(value: unknown): string {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  if (typeof value === "number") return value.toLocaleString("en-IE");
  if (typeof value === "string") {
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(value)) {
      const date = new Date(value);
      if (!Number.isNaN(date.getTime())) {
        return new Intl.DateTimeFormat("en-IE", {
          day: "2-digit",
          month: "short",
          year: "numeric",
        }).format(date);
      }
    }
    return value.length > 80 ? `${value.slice(0, 77)}…` : value;
  }
  return JSON.stringify(value);
}

function RecordPreviewTable({
  records,
  columns,
  entity,
  maxColumns = 6,
}: {
  records: Array<Record<string, unknown>>;
  columns: DataverseColumnMetadata[];
  entity: string;
  maxColumns?: number;
}) {
  const keys: string[] = [];
  for (const record of records) {
    for (const key of Object.keys(record)) {
      if (key.startsWith("@") || key.includes("@")) continue;
      if (!keys.includes(key)) keys.push(key);
    }
  }
  const logicalNameFor = (key: string) =>
    key.startsWith("_") && key.endsWith("_value") ? key.slice(1, -6) : key;
  const labelFor = (key: string) => {
    const logicalName = logicalNameFor(key);
    return (
      columns.find((column) => column.logicalName === logicalName)?.displayName ??
      logicalName
    );
  };
  const primaryId = `${entity}id`;
  const score = (key: string) => {
    if (key === primaryId) return 3;
    if (key.startsWith("_") && key.endsWith("_value")) return 2;
    if (records.every((record) => isGuidLike(record[key]))) return 2;
    if (["name", "fullname", "title", "subject"].includes(key)) return 0;
    return 1;
  };
  const ranked = [...keys].sort(
    (a, b) => score(a) - score(b) || keys.indexOf(a) - keys.indexOf(b),
  );
  const visible = ranked.slice(0, maxColumns);
  const hiddenCount = keys.length - visible.length;
  const rowKey = (record: Record<string, unknown>, index: number) =>
    String(record[primaryId] ?? record.id ?? index);
  const cellValue = (record: Record<string, unknown>, key: string) => {
    const formatted = record[`${key}@OData.Community.Display.V1.FormattedValue`];
    return formatPreviewValue(formatted ?? record[key]);
  };
  const surfaceRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    surfaceRef.current?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }, [records]);
  return (
    <div className="preview-table-surface" ref={surfaceRef}>
      <div className="preview-table-heading">
        <b>
          {records.length} matching record{records.length === 1 ? "" : "s"} loaded
        </b>
        {hiddenCount > 0 && (
          <small>
            Showing {visible.length} of {keys.length} columns
          </small>
        )}
      </div>
      <Table aria-label="Record preview" className="preview-table" size="small">
        <TableHeader>
          <TableRow>
            {visible.map((key) => (
              <TableHeaderCell key={key} title={key}>
                {labelFor(key)}
              </TableHeaderCell>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {records.map((record, index) => (
            <TableRow key={rowKey(record, index)}>
              {visible.map((key) => (
                <TableCell key={key} title={String(record[key] ?? "")}>
                  <span className={isGuidLike(record[key]) ? "preview-guid" : undefined}>
                    {cellValue(record, key)}
                  </span>
                </TableCell>
              ))}
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
}

function CompiledRecipe({
  step,
  draft,
  title = "Recipe v3",
}: {
  step: number;
  draft: BuilderDraft;
  title?: string;
}) {
  const status = [
    "context ready",
    "prompt bound",
    "destination valid",
    "flow prepared",
    "ready to publish",
  ][step];
  const [showJson, setShowJson] = useState(false);
  const recipe = compileRecipe(draft);
  const estimatedTokens = estimateRecipeTokens({
    promptContent: draft.promptContent,
    sourceFields: draft.sourceFields,
    relationships: draft.relationships,
  });
  const assets = [
    "csp_aisummaryconfig record",
    draft.promptId ? "AI Prompt binding" : null,
    "Flow definition",
  ].filter(Boolean);
  return (
    <aside className="compiled-panel">
      <div className="compiled-head">
        <div>
          <p className="micro-label text-brand-cyan">Compiled view</p>
          <h2>{title}</h2>
        </div>
        <span>
          <i />
          {status}
        </span>
      </div>
      <div className="recipe-line">
        <span>OBJECTIVE</span>
        <b>{draft.name}</b>
      </div>
      <div className="recipe-code">
        <div>
          <span>01</span>
          <code>source</code>
          <b>table · {draft.entity}</b>
        </div>
        <div>
          <span>02</span>
          <code>fields</code>
          <b>{draft.sourceFields.length} selected</b>
        </div>
        <div>
          <span>03</span>
          <code>relation</code>
          <b>{String(draft.relationships[0]?.entity ?? "none")}</b>
        </div>
        <div>
          <span>04</span>
          <code>prompt</code>
          <b>{draft.promptKey}</b>
        </div>
        <div>
          <span>05</span>
          <code>model</code>
          <b>{draft.model}</b>
        </div>
        <div>
          <span>06</span>
          <code>output</code>
          <b>
            {draft.outputEntity}.{draft.outputField}
          </b>
        </div>
        <div>
          <span>07</span>
          <code>trigger</code>
          <b>{draft.triggerType}</b>
        </div>
      </div>
      <div className="compile-summary">
        <p>
          <span>Estimated context</span>
          <b>~{estimatedTokens.toLocaleString("en-US")} tokens</b>
        </p>
        <p>
          <span>Flow pattern</span>
          <b>
            {
              triggerPresentation(draft.triggerType, draft.triggerColumns)
                .pattern
            }
          </b>
        </p>
        <p>
          <span>Components</span>
          <b>
            {assets.length} asset{assets.length === 1 ? "" : "s"}
          </b>
        </p>
      </div>
      <div className="compiled-footer">
        <Code />
        <span>
          <b>Valid JSON</b>
          <small>summary-recipe.schema.json</small>
        </span>
        <button
          type="button"
          aria-expanded={showJson}
          onClick={() => setShowJson((value) => !value)}
        >
          {showJson ? "Hide JSON" : "View JSON"}
        </button>
      </div>
      {showJson && (
        <pre className="recipe-json" aria-label="Compiled recipe JSON">
          {JSON.stringify(recipe, null, 2)}
        </pre>
      )}
    </aside>
  );
}

const paneContent: Record<
  PaneKind,
  {
    title: string;
    eyebrow: string;
    value: string;
    technical: string;
    options: string[];
  }
> = {
  table: {
    title: "Primary table",
    eyebrow: "Summary source",
    value: "Account · account",
    technical: "Dataverse table",
    options: [
      "Account · account",
      "Case · incident",
      "Opportunity · opportunity",
    ],
  },
  execution: {
    title: "Summary mode",
    eyebrow: "Summary pattern",
    value: "One summary per record",
    technical: "PerRecord",
    options: [
      "One summary per record",
      "Consolidated summary",
      "Modified records only",
    ],
  },
  prompt: {
    title: "AI Prompt",
    eyebrow: "Reusable asset",
    value: "Executive summary prompt",
    technical: "prompt_account_brief_v3",
    options: [
      "Executive summary prompt",
      "Management brief",
      "Commercial risk summary",
    ],
  },
  model: {
    title: "Model",
    eyebrow: "Generation",
    value: "GPT-4.1 mini",
    technical: "AI Builder model",
    options: ["GPT-4.1 mini", "GPT-4.1", "GPT-4o mini"],
  },
  fields: {
    title: "Selected fields",
    eyebrow: "Summary context",
    value: "Account fields",
    technical: "Dataverse metadata",
    options: [],
  },
  records: {
    title: "Records to process",
    eyebrow: "Record selection",
    value: "Custom FetchXML",
    technical: "csp_fetchxml",
    options: [],
  },
  relationships: {
    title: "Related records",
    eyebrow: "Context source",
    value: "Open and recent activities",
    technical: "csp_relationships",
    options: ["Open and recent activities", "No related records"],
  },
  contextQuery: {
    title: "Context query",
    eyebrow: "Per-record context",
    value: "FetchXML per record",
    technical: "csp_relatedfetchxml",
    options: [],
  },
  outputTable: {
    title: "Destination table",
    eyebrow: "Summary destination",
    value: "Account · account",
    technical: "Dataverse table",
    options: ["Account · account", "Summary cache · csp_aisummarycache"],
  },
  outputField: {
    title: "Destination column",
    eyebrow: "Summary destination",
    value: "AI summary · csp_aisummary",
    technical: "Dataverse column",
    options: [
      "AI summary · csp_aisummary",
      "Description · description",
      "Name · name",
    ],
  },
};

function TableCatalogPane({
  purpose,
  current,
  onClose,
  onApply,
}: {
  purpose: "source" | "destination";
  current: string;
  onClose: () => void;
  onApply: (table: DataverseTableMetadata) => void;
}) {
  const { data: tables = [], isLoading } = useDataverseTables();
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(current);
  const visible = tables.filter((table) =>
    `${table.displayName} ${table.logicalName}`
      .toLowerCase()
      .includes(search.toLowerCase()),
  );
  const selectedTable = tables.find((table) => table.logicalName === selected);
  const selectedMetadata = useDataverseTableMetadata(selectedTable);
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside className="edit-pane" aria-label={`Choose ${purpose} table`}>
        <header>
          <div>
            <p className="micro-label text-brand-blue">Dataverse metadata</p>
            <h2>
              {purpose === "source" ? "Source table" : "Destination table"}
            </h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <div className="pane-current">
          <span>
            <Database />
          </span>
          <div>
            <small>Current table</small>
            <b>
              {tables.find((table) => table.logicalName === current)
                ?.displayName ?? current}
            </b>
            <code>{current}</code>
          </div>
        </div>
        <section>
          <label className="studio-field">
            <span>Find a table</span>
            <Input
              appearance="outline"
              className="w-full"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search display or logical name"
            />
          </label>
          <p className="field-caption mt-4">Available tables</p>
          {isLoading ? (
            <p className="text-sm text-muted-foreground">
              Loading Dataverse tables…
            </p>
          ) : (
            visible.map((table) => (
              <button
                type="button"
                key={table.logicalName}
                aria-label={`${table.displayName} · ${table.logicalName}`}
                className={selected === table.logicalName ? "selected" : ""}
                onClick={() => setSelected(table.logicalName)}
              >
                <span>{selected === table.logicalName ? <Check /> : null}</span>
                <div>
                  <b>{table.displayName}</b>
                  <code>
                    {table.logicalName} · {table.entitySetName}
                  </code>
                </div>
              </button>
            ))
          )}
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            disabled={!selectedMetadata.data || selectedMetadata.isFetching}
            onClick={() =>
              selectedMetadata.data && onApply(selectedMetadata.data)
            }
          >
            {selectedMetadata.isFetching ? "Reading metadata…" : "Apply table"}
          </Button>
        </footer>
      </aside>
    </>
  );
}

function FieldsPane({
  entity,
  entitySetName,
  selectedFields,
  onClose,
  onApply,
}: {
  entity: string;
  entitySetName: string;
  selectedFields: string[];
  onClose: () => void;
  onApply: (fields: string[]) => void;
}) {
  const { data: columns = [], isLoading } = useDataverseColumns(
    entity,
    entitySetName,
  );
  const [selected, setSelected] = useState(selectedFields);
  const toggle = (field: string) =>
    setSelected((current) =>
      current.includes(field)
        ? current.filter((item) => item !== field)
        : [...current, field],
    );
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside className="edit-pane" aria-label="Choose source columns">
        <header>
          <div>
            <p className="micro-label text-brand-blue">Dataverse metadata</p>
            <h2>Source columns</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <div className="pane-current">
          <span>
            <Database />
          </span>
          <div>
            <small>Current selection</small>
            <b>{selected.length} columns</b>
            <code>{entity}</code>
          </div>
        </div>
        <section>
          <p className="field-caption">Readable columns</p>
          {isLoading ? (
            <p className="text-sm text-muted-foreground">Loading columns…</p>
          ) : (
            columns
              .filter((column) => column.readable)
              .map((column) => {
                const active = selected.includes(column.logicalName);
                return (
                  <button
                    type="button"
                    key={column.logicalName}
                    aria-label={`${column.displayName} · ${column.logicalName}`}
                    aria-pressed={active}
                    onClick={() => toggle(column.logicalName)}
                    className={active ? "selected" : ""}
                  >
                    <span>{active ? <Check /> : null}</span>
                    <div>
                      <b>{column.displayName}</b>
                      <code>
                        {column.logicalName} · {column.type}
                      </code>
                    </div>
                  </button>
                );
              })
          )}
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            onClick={() => onApply(selected)}
            disabled={selected.length === 0}
          >
            Apply changes
          </Button>
        </footer>
      </aside>
    </>
  );
}

function OutputFieldPane({
  entity,
  entitySetName,
  current,
  onClose,
  onApply,
}: {
  entity: string;
  entitySetName: string;
  current: string;
  onClose: () => void;
  onApply: (field: string) => void;
}) {
  const { data: columns = [], isLoading } = useDataverseColumns(
    entity,
    entitySetName,
  );
  const [selected, setSelected] = useState(current);
  const compatible = columns.filter(
    (column) => column.writable && ["String", "Memo"].includes(column.type),
  );
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside className="edit-pane" aria-label="Choose destination column">
        <header>
          <div>
            <p className="micro-label text-brand-blue">
              Writable Dataverse columns
            </p>
            <h2>Destination column</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <div className="pane-current">
          <span>
            <Target />
          </span>
          <div>
            <small>Current column</small>
            <b>{current}</b>
            <code>
              {entity}.{current}
            </code>
          </div>
        </div>
        <section>
          <p className="field-caption">Text and multiline text</p>
          {isLoading ? (
            <p className="text-sm text-muted-foreground">Loading columns…</p>
          ) : (
            compatible.map((column) => (
              <button
                type="button"
                key={column.logicalName}
                aria-label={`${column.displayName} · ${column.logicalName}`}
                className={selected === column.logicalName ? "selected" : ""}
                onClick={() => setSelected(column.logicalName)}
              >
                <span>
                  {selected === column.logicalName ? <Check /> : null}
                </span>
                <div>
                  <b>{column.displayName}</b>
                  <code>
                    {column.logicalName} · {column.type}
                  </code>
                </div>
              </button>
            ))
          )}
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button disabled={!selected} onClick={() => onApply(selected)}>
            Apply column
          </Button>
        </footer>
      </aside>
    </>
  );
}

type QueryDesignerMode = "filters" | "view" | "fetchxml";

function relatedSelectionFetchXml(relationship?: Record<string, unknown>) {
  if (relationship?.fetchXml) return String(relationship.fetchXml);
  const days = Number(relationship?.windowDays ?? 30);
  const entity = String(relationship?.entity ?? "activitypointer");
  return `<fetch top="12">\n  <entity name="${entity}">\n    <all-attributes />\n    <filter type="and">\n      <condition attribute="createdon" operator="last-x-days" value="${days}" />\n    </filter>\n    <order attribute="createdon" descending="true" />\n  </entity>\n</fetch>`;
}

type QueryDesignerResult = {
  fetchXml: string;
  maxRecords: number;
  queryMode: number;
  viewId?: string;
  viewType?: "system" | "personal";
  relatedTable?: DataverseTableMetadata;
  relationship?: DataverseRelationshipMetadata;
  fields?: string[];
};

function QueryDesignerPane({
  purpose,
  targetEntity,
  entitySetName,
  sourceEntity,
  sourceEntitySetName,
  relationship: initialRelationship,
  queryMode = 100000002,
  initialViewId,
  initialViewType,
  fetchXml,
  maxRecords,
  onClose,
  onApply,
}: {
  purpose: "records" | "related";
  targetEntity: string;
  entitySetName: string;
  sourceEntity?: string;
  sourceEntitySetName?: string;
  relationship?: Record<string, unknown>;
  queryMode?: number;
  initialViewId?: string;
  initialViewType?: "system" | "personal";
  fetchXml: string;
  maxRecords: number;
  onClose: () => void;
  onApply: (value: QueryDesignerResult) => void;
}) {
  const [mode, setMode] = useState<QueryDesignerMode>(
    queryMode === 100000000 || queryMode === 100000001 ? "view" : "filters",
  );
  const [query, setQuery] = useState(fetchXml);
  const [limit, setLimit] = useState(maxRecords);
  const [viewId, setViewId] = useState(initialViewId ?? "");
  const [viewType, setViewType] = useState<"system" | "personal">(
    initialViewType ?? (queryMode === 100000001 ? "personal" : "system"),
  );
  const [relationshipKey, setRelationshipKey] = useState(
    String(initialRelationship?.schemaName ?? ""),
  );
  const [relatedFields, setRelatedFields] = useState<string[]>(
    Array.isArray(initialRelationship?.fields)
      ? initialRelationship.fields.filter(
          (field): field is string => typeof field === "string",
        )
      : [],
  );
  const tables = useDataverseTables();
  const relationships = useDataverseRelationships(
    sourceEntity ?? targetEntity,
    sourceEntitySetName ?? entitySetName,
  );
  const availableRelationships = relationships.data ?? [];
  const selectedRelationship =
    availableRelationships.find(
      (item) => item.schemaName === relationshipKey,
    ) ?? availableRelationships.find((item) => item.entity === targetEntity);
  const effectiveEntity =
    purpose === "related"
      ? (selectedRelationship?.entity ?? targetEntity)
      : targetEntity;
  const effectiveTable = tables.data?.find(
    (table) => table.logicalName === effectiveEntity,
  );
  const effectiveEntitySet = effectiveTable?.entitySetName ?? entitySetName;
  const relatedColumns = useDataverseColumns(
    effectiveEntity,
    effectiveEntitySet,
  );
  const systemViews = useSystemViews(effectiveEntity);
  const userViews = useUserViews(effectiveEntity);
  const preview = useDataverseRecordPreview();
  const availableViews =
    viewType === "system" ? (systemViews.data ?? []) : (userViews.data ?? []);
  const title = purpose === "records" ? "Record selection" : "Related records";
  const technical =
    purpose === "records" ? "csp_fetchxml" : "csp_relatedfetchxml";
  const chooseView = (id: string) => {
    setViewId(id);
    const selected = availableViews.find((view) => view.id === id);
    if (selected?.fetchXml) setQuery(formatFetchXml(selected.fetchXml));
  };
  const defaultRelatedFields = (relatedColumns.data ?? [])
    .filter((column) => column.readable && !column.logicalName.endsWith("id"))
    .slice(0, 3)
    .map((column) => column.logicalName);
  const effectiveRelatedFields = relatedFields.length
    ? relatedFields
    : defaultRelatedFields;
  const changeRelatedTable = (schemaName: string) => {
    const nextRelationship = availableRelationships.find(
      (item) => item.schemaName === schemaName,
    );
    if (!nextRelationship) return;
    setRelationshipKey(schemaName);
    setRelatedFields([]);
    setViewId("");
    setQuery(
      relatedSelectionFetchXml({
        entity: nextRelationship.entity,
        maxRecords: limit,
      }),
    );
  };
  const toggleRelatedField = (field: string) =>
    setRelatedFields((current) => {
      const base = current.length ? current : defaultRelatedFields;
      return base.includes(field)
        ? base.filter((item) => item !== field)
        : [...base, field];
    });
  const apply = () =>
    onApply({
      fetchXml: query.trim(),
      maxRecords: limit,
      queryMode:
        mode === "view"
          ? viewType === "system"
            ? 100000000
            : 100000001
          : 100000002,
      viewId: mode === "view" ? viewId : undefined,
      viewType: mode === "view" ? viewType : undefined,
      ...(purpose === "related"
        ? {
            relatedTable: effectiveTable,
            relationship: selectedRelationship,
            fields: effectiveRelatedFields,
          }
        : {}),
    });

  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside
        className="edit-pane query-designer-pane"
        aria-label={`Configure ${title.toLowerCase()}`}
      >
        <header>
          <div>
            <p className="micro-label text-brand-blue">
              {purpose === "records"
                ? "Rows entering the automation"
                : "Rows enriching each summary"}
            </p>
            <h2>{title}</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        {purpose === "related" && (
          <div className="space-y-3 border-b border-border px-5 py-4">
            <label className="studio-field">
              <span>Related Dataverse table</span>
              <select
                aria-label="Related Dataverse table"
                value={selectedRelationship?.schemaName ?? ""}
                onChange={(event) => changeRelatedTable(event.target.value)}
              >
                {availableRelationships.map((item) => {
                  const table = tables.data?.find(
                    (candidate) => candidate.logicalName === item.entity,
                  );
                  return (
                    <option
                      key={item.schemaName || item.entity}
                      value={item.schemaName}
                    >
                      {table?.displayName ?? item.entity} · {item.entity} ·{" "}
                      {item.schemaName}
                    </option>
                  );
                })}
              </select>
            </label>
            <div>
              <p className="field-caption">Columns passed to the prompt</p>
              <div className="mt-2 flex flex-wrap gap-2">
                {relatedColumns.isLoading ? (
                  <small>Loading columns…</small>
                ) : (
                  (relatedColumns.data ?? [])
                    .filter((column) => column.readable)
                    .map((column) => {
                      const active = effectiveRelatedFields.includes(
                        column.logicalName,
                      );
                      return (
                        <button
                          type="button"
                          key={column.logicalName}
                          aria-label={`${column.displayName} · ${column.logicalName}`}
                          aria-pressed={active}
                          className={cn(
                            "rounded-md border px-2.5 py-1.5 text-left text-xs",
                            active
                              ? "border-cyan-500 bg-cyan-50 text-slate-950"
                              : "border-border bg-white",
                          )}
                          onClick={() => toggleRelatedField(column.logicalName)}
                        >
                          {column.displayName}
                          <code className="ml-1 text-[10px] text-muted-foreground">
                            {column.logicalName}
                          </code>
                        </button>
                      );
                    })
                )}
              </div>
            </div>
          </div>
        )}
        <div
          className="query-mode-tabs"
          role="tablist"
          aria-label={`${title} method`}
        >
          <button
            type="button"
            role="tab"
            aria-selected={mode === "filters"}
            onClick={() => setMode("filters")}
          >
            <Settings />
            Build filters
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mode === "view"}
            onClick={() => setMode("view")}
          >
            <Grid />
            Choose a view
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mode === "fetchxml"}
            onClick={() => {
              setMode("fetchxml");
              setQuery((current) => formatFetchXml(current));
            }}
          >
            <Code />
            Edit FetchXML
          </button>
        </div>
        <section className="query-designer-body">
          {mode === "filters" && (
            <>
              <div className="query-mode-intro">
                <b>Build it like a Dataverse view</b>
                <small>
                  Add conditions, sorting, and a record limit. The app compiles
                  valid FetchXML for the flow.
                </small>
              </div>
              <VisualFilterBuilder
                targetEntity={effectiveEntity}
                onGenerate={(generated) => setQuery(generated)}
              />
            </>
          )}
          {mode === "view" && (
            <>
              <div className="query-mode-intro">
                <b>Reuse an existing Dataverse view</b>
                <small>
                  Select a system or personal view. Its FetchXML becomes part of
                  this configuration.
                </small>
              </div>
              <div className="view-type-toggle">
                <button
                  type="button"
                  className={viewType === "system" ? "active" : ""}
                  onClick={() => {
                    setViewType("system");
                    setViewId("");
                  }}
                >
                  System views
                </button>
                <button
                  type="button"
                  className={viewType === "personal" ? "active" : ""}
                  onClick={() => {
                    setViewType("personal");
                    setViewId("");
                  }}
                >
                  Personal views
                </button>
              </div>
              <label className="studio-field">
                <span>Saved view</span>
                <select
                  aria-label="Saved view"
                  value={viewId}
                  onChange={(event) => chooseView(event.target.value)}
                >
                  <option value="">Select a {viewType} view…</option>
                  {availableViews.map((view) => (
                    <option key={view.id} value={view.id}>
                      {view.name}
                    </option>
                  ))}
                </select>
              </label>
              <p className="query-source-status">
                {(
                  viewType === "system"
                    ? systemViews.isFetching
                    : userViews.isFetching
                )
                  ? "Loading Dataverse views…"
                  : `${availableViews.length} ${viewType} views available for ${effectiveEntity}`}
              </p>
            </>
          )}
          {mode === "fetchxml" && (
            <>
              <div className="query-mode-intro">
                <b>Advanced FetchXML</b>
                <small>
                  Edit the compiled query directly for joins, nested groups, or
                  operators not exposed by the filter builder.
                </small>
              </div>
              <label className="studio-field">
                <span>
                  {purpose === "records"
                    ? "Record selection FetchXML"
                    : "Related-record FetchXML"}
                </span>
                <textarea
                  aria-label={
                    purpose === "records"
                      ? "Record selection FetchXML"
                      : "Related-record FetchXML"
                  }
                  className="query-code-editor"
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  spellCheck={false}
                />
              </label>
            </>
          )}
          <div className="query-output">
            <span>
              <Code />
            </span>
            <div>
              <small>Output stored in</small>
              <b>{technical}</b>
            </div>
            <code>{query.length.toLocaleString("en-US")} characters</code>
          </div>
          {preview.data && (
            <RecordPreviewTable
              records={preview.data}
              columns={relatedColumns.data ?? []}
              entity={effectiveEntity}
            />
          )}
          {preview.isError && (
            <p className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
              {preview.error instanceof Error
                ? preview.error.message
                : "Dataverse could not execute this query."}
            </p>
          )}
        </section>
        <footer>
          <div className="query-limit">
            <label>
              Maximum records{" "}
              <input
                type="number"
                min={1}
                max={5000}
                value={limit}
                onChange={(event) =>
                  setLimit(
                    Math.max(
                      1,
                      Math.min(5000, Number(event.target.value) || 1),
                    ),
                  )
                }
              />
            </label>
          </div>
          <Button
            type="button"
            variant="outline"
            aria-label="Preview matching records"
            disabled={!query.trim().startsWith("<fetch") || preview.isPending}
            onClick={() =>
              preview.mutate({
                logicalName: effectiveEntity,
                entitySetName: effectiveEntitySet,
                fetchXml: query,
                limit: Math.min(limit, 10),
              })
            }
          >
            {preview.isPending ? "Loading…" : "Preview records"}
          </Button>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            onClick={apply}
            disabled={
              !query.trim().startsWith("<fetch") ||
              (mode === "view" && !viewId) ||
              (purpose === "related" && !selectedRelationship)
            }
          >
            Apply selection
          </Button>
        </footer>
      </aside>
    </>
  );
}

function ContextQueryPane({
  fetchXml,
  onClose,
  onApply,
}: {
  fetchXml: string;
  onClose: () => void;
  onApply: (fetchXml: string) => void;
}) {
  const [query, setQuery] = useState(fetchXml);
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside className="edit-pane" aria-label="Edit context query">
        <header>
          <div>
            <p className="micro-label text-brand-blue">Per-record context</p>
            <h2>Context FetchXML</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <div className="pane-current">
          <span>
            <Code />
          </span>
          <div>
            <small>Configuration column</small>
            <b>Context query</b>
            <code>csp_relatedfetchxml</code>
          </div>
        </div>
        <section className="space-y-4">
          <div className="query-guidance">
            <GitBranch />
            <p>
              <b>Runs once for every selected record</b>
              <small>
                Use <code>{"{{recordId}}"}</code> where the generated flow must
                inject the current Dataverse row identifier.
              </small>
            </p>
          </div>
          <label className="studio-field">
            <span>Context FetchXML</span>
            <textarea
              aria-label="Context FetchXML"
              className="min-h-80 w-full resize-none rounded-lg border border-border bg-slate-950 p-3 font-mono text-xs leading-5 text-cyan-100 outline-none focus:border-brand-cyan"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              spellCheck={false}
            />
          </label>
          <p className="text-xs leading-5 text-muted-foreground">
            This query builds the payload sent to the AI Prompt. It can include
            source fields, linked entities, filters, ordering, and record
            limits.
          </p>
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            onClick={() => onApply(query.trim())}
            disabled={
              !query.trim().startsWith("<fetch") ||
              !query.includes("{{recordId}}")
            }
          >
            Apply changes
          </Button>
        </footer>
      </aside>
    </>
  );
}

function TestRecordPane({
  draft,
  onClose,
}: {
  draft: BuilderDraft;
  onClose: () => void;
}) {
  const records = useDataverseRecordPreview();
  const context = useDataverseRecordPreview();
  const [selectedId, setSelectedId] = useState("");
  const idField =
    draft.entityIdField || getTargetEntityIdentity(draft.entity).idField;
  const entitySetName =
    draft.entitySetName || getTargetEntityIdentity(draft.entity).entitySet;
  const labelFor = (record: Record<string, unknown>) =>
    String(
      record.name ??
        record.fullname ??
        record.title ??
        record.subject ??
        record[idField] ??
        "Dataverse record",
    );
  const compile = () => {
    if (!selectedId) return;
    context.mutate({
      logicalName: draft.entity,
      entitySetName,
      fetchXml: draft.relatedFetchXml.replaceAll("{{recordId}}", selectedId),
      limit: 10,
    });
  };
  const selectableRecords = (records.data ?? []).filter((record) =>
    Boolean(record[idField]),
  );
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside
        className="edit-pane query-designer-pane"
        aria-label="Test with a Dataverse record"
      >
        <header>
          <div>
            <p className="micro-label text-brand-blue">Live Dataverse test</p>
            <h2>Test with a record</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <section className="query-designer-body">
          <div className="query-guidance">
            <Database />
            <p>
              <b>Use the current record selection</b>
              <small>
                The app executes <code>csp_fetchxml</code>, then compiles the
                per-record context exactly as the generated flow will.
              </small>
            </p>
          </div>
          <Button
            type="button"
            variant="outline"
            aria-label="Load matching records"
            disabled={records.isPending}
            onClick={() =>
              records.mutate({
                logicalName: draft.entity,
                entitySetName,
                fetchXml: draft.fetchXml,
                limit: 10,
              })
            }
          >
            {records.isPending ? "Loading records…" : "Load matching records"}
          </Button>
          {records.data && (
            <div className="mt-4 space-y-2">
              <p className="field-caption">Choose a record</p>
              {selectableRecords.map((record) => {
                const id = String(record[idField]);
                return (
                  <button
                    type="button"
                    key={id}
                    className={cn(
                      "flex w-full items-center gap-3 rounded-lg border p-3 text-left",
                      selectedId === id && "border-cyan-500 bg-cyan-50",
                    )}
                    onClick={() => setSelectedId(id)}
                  >
                    <span className="radio-dot" />
                    <span>
                      <b className="block text-sm">{labelFor(record)}</b>
                      <code className="text-xs">{id}</code>
                    </span>
                  </button>
                );
              })}
              {selectableRecords.length === 0 && (
                <p className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm text-amber-900">
                  The current query must include <code>{idField}</code> to test
                  a record.
                </p>
              )}
            </div>
          )}
          {context.data && (
            <div className="mt-4 rounded-xl bg-slate-950 p-4 text-cyan-100">
              <p className="micro-label text-brand-cyan">
                Compiled prompt context
              </p>
              <pre className="mt-3 max-h-72 overflow-auto whitespace-pre-wrap text-xs">
                {JSON.stringify(context.data, null, 2)}
              </pre>
            </div>
          )}
          {(records.isError || context.isError) && (
            <p className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
              {[records.error, context.error]
                .map((error) => (error instanceof Error ? error.message : ""))
                .filter(Boolean)
                .join(" ") || "Dataverse could not execute the current recipe."}
            </p>
          )}
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Close
          </Button>
          <Button onClick={compile} disabled={!selectedId || context.isPending}>
            {context.isPending ? "Compiling…" : "Compile context"}
          </Button>
        </footer>
      </aside>
    </>
  );
}

function EditPane({
  pane,
  current,
  options,
  onClose,
  onApply,
}: {
  pane: PaneKind;
  current: string;
  options: string[];
  onClose: () => void;
  onApply: (value: string) => void;
}) {
  const content = paneContent[pane];
  const [selected, setSelected] = useState(current);
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside
        className="edit-pane"
        aria-label={`Edit ${content.title.toLowerCase()}`}
      >
        <header>
          <div>
            <p className="micro-label text-brand-blue">{content.eyebrow}</p>
            <h2>{content.title}</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        <div className="pane-current">
          <span>
            <CheckCircle />
          </span>
          <div>
            <small>Current value</small>
            <b>{current}</b>
            <code>{content.technical}</code>
          </div>
        </div>
        <section>
          <p className="field-caption">Select value</p>
          {options.map((option) => (
            <button
              type="button"
              key={option}
              onClick={() => setSelected(option)}
              className={selected === option ? "selected" : ""}
            >
              <span>{selected === option ? <Check /> : null}</span>
              {option}
            </button>
          ))}
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button onClick={() => onApply(selected)}>Apply changes</Button>
        </footer>
      </aside>
    </>
  );
}

function Builder() {
  const { data } = useStudioRuntime();
  const { configurationId } = useParams();
  const isNew = !configurationId;
  const createConfiguration = useCreateStudioConfiguration();
  const updateConfiguration = useUpdateStudioConfiguration();
  const updatePrompt = useUpdateStudioPrompt();
  const configuration =
    configurationId === "account-operations"
      ? data?.configurations[0]
      : data?.configurations.find((item) => item.id === configurationId);
  const prompts = useMemo(() => data?.prompts ?? [], [data?.prompts]);
  const location = useLocation();
  const navigationState = (location.state ?? {}) as {
    openRecipe?: boolean;
    openTest?: boolean;
  };
  const [step, setStep] = useState(0);
  const [recipeOpen, setRecipeOpen] = useState(
    Boolean(navigationState.openRecipe),
  );
  const [testRecordOpen, setTestRecordOpen] = useState(
    Boolean(navigationState.openTest),
  );
  const [pane, setPane] = useState<PaneKind | null>(null);
  const [validationErrors, setValidationErrors] = useState<string[]>([]);
  const [draft, setDraft] = useState<BuilderDraft>(() =>
    draftFromConfiguration(),
  );
  const navigate = useNavigate();
  // Dataverse arrives asynchronously after the local-first editor has mounted.
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect
    if (configuration) setDraft(draftFromConfiguration(configuration, prompts));
    else if (isNew && prompts.length)
      setDraft((current) =>
        current.promptId ? current : draftFromConfiguration(undefined, prompts),
      );
  }, [configuration, isNew, prompts]);
  const notify = useStudioToast();
  const save = async (next: BuilderDraft, changes: Record<string, unknown>) => {
    setDraft(next);
    setPane(null);
    setValidationErrors([]);
    try {
      if (configuration) {
        await updateConfiguration.mutateAsync({
          id: configuration.id,
          changes: {
            ...changes,
            csp_configurationjson: JSON.stringify(compileRecipe(next)),
          },
        });
        notify("success", "Configuration saved", `Changes stored on "${configuration.name}".`);
        return configuration.id;
      }
      if (isNew) {
        const id = await createConfiguration.mutateAsync(
          buildDraftConfigurationPayload(next),
        );
        notify("success", "Draft created", "The configuration was created in Dataverse.");
        navigate(`/configurations/${id}`, { replace: true });
        return id;
      }
    } catch (error) {
      notify("error", "Could not save the configuration", errorMessage(error, ""));
    }
    return undefined;
  };
  const paneOptions =
    pane === "prompt"
      ? prompts.map((item) => item.name)
      : pane
        ? paneContent[pane].options
        : [];
  const paneCurrent =
    pane === "table"
      ? ({
          account: "Account · account",
          incident: "Case · incident",
          opportunity: "Opportunity · opportunity",
        }[draft.entity] ?? draft.entity)
      : pane === "execution"
        ? draft.mode === 100000000
          ? "One summary per record"
          : "Consolidated summary"
        : pane === "prompt"
          ? draft.promptName
          : pane === "outputTable"
            ? draft.outputEntity === "account"
              ? "Account · account"
              : "Summary cache · csp_aisummarycache"
            : pane === "outputField"
              ? ({
                  csp_aisummary: "AI summary · csp_aisummary",
                  description: "Description · description",
                  name: "Name · name",
                }[draft.outputField] ?? draft.outputField)
              : pane === "records"
                ? "Custom FetchXML"
                : pane === "relationships"
                  ? draft.relationships.length
                    ? "Open and recent activities"
                    : "No related records"
                  : pane === "contextQuery"
                    ? "FetchXML per record"
                    : draft.model
                        .replace("gpt-", "GPT-")
                        .replace("-mini", " mini");
  const applyPane = async (value: string) => {
    if (!pane) return;
    if (pane === "execution") {
      const mode = value === "Consolidated summary" ? 100000001 : 100000000;
      await save({ ...draft, mode }, { csp_mode: mode });
    }
    if (pane === "model") {
      const model = value.toLowerCase().replace(" ", "-");
      await save({ ...draft, model }, { csp_model: model });
    }
    if (pane === "prompt") {
      const prompt = prompts.find((item) => item.name === value);
      if (prompt)
        await save(
          {
            ...draft,
            promptId: prompt.id,
            promptName: prompt.name,
            promptKey: prompt.key,
            promptContent: prompt.content,
          },
          { "csp_Prompt@odata.bind": `/csp_aiprompts(${prompt.id})` },
        );
    }
  };
  const updateDraft = async (changes: Partial<BuilderDraft>) => {
    const next = { ...draft, ...changes };
    const fields: Record<string, unknown> = {};
    if (changes.sourceFields || changes.relationships) {
      next.relatedFetchXml = buildContextFetchXml(
        next.entity,
        next.sourceFields,
        next.relationships,
        next.entityIdField,
      );
      fields.csp_sourcefields = JSON.stringify(next.sourceFields);
      fields.csp_relationships = JSON.stringify(next.relationships);
      fields.csp_relatedfetchxml = next.relatedFetchXml;
    }
    if (changes.sourceFields && draft.triggerColumns.length === 0) {
      next.triggerColumns = changes.sourceFields.filter(
        (field) => field !== next.outputField,
      );
      fields.csp_triggercolumns = JSON.stringify(next.triggerColumns);
    }
    if (changes.triggerColumns)
      fields.csp_triggercolumns = JSON.stringify(changes.triggerColumns);
    if (changes.outputEntity) fields.csp_outputentity = changes.outputEntity;
    if (changes.outputField) fields.csp_outputfield = changes.outputField;
    await save(next, fields);
  };
  const savePromptContent = async (content: string) => {
    if (!draft.promptId) return;
    await updatePrompt.mutateAsync({ id: draft.promptId, content });
    setDraft((current) => ({ ...current, promptContent: content }));
    notify("success", "Prompt saved", `"${draft.promptName}" was updated in Dataverse.`);
  };
  const publish = async () => {
    const errors = validateSummaryConfiguration(draft);
    setValidationErrors(errors);
    if (errors.length) {
      setStep(4);
      notify("warning", "Resolve validation items before publishing", errors[0]);
      return;
    }
    try {
      let id = configuration?.id;
      if (configuration) {
        await updateConfiguration.mutateAsync({
          id: configuration.id,
          changes: buildConfigurationPayload(draft),
        });
      } else if (isNew) {
        id = await createConfiguration.mutateAsync(
          buildConfigurationPayload(draft),
        );
      }
      if (!id) return;
      await updateConfiguration.mutateAsync({
        id,
        changes: buildPublishSignal(),
      });
      notify("success", "Configuration published", "The recipe is ready for the generator backend.");
      navigate(`/published/${id}`);
    } catch (error) {
      notify("error", "Could not publish the configuration", errorMessage(error, ""));
    }
  };
  const content = [
    <DataStep onEdit={setPane} draft={draft} accounts={data?.accounts ?? []} />,
    <PromptStep
      key={draft.promptId}
      onEdit={setPane}
      draft={draft}
      onSavePrompt={savePromptContent}
      savingPrompt={updatePrompt.isPending}
    />,
    <DestinationStep draft={draft} onEdit={setPane} onChange={updateDraft} />,
    <ExecutionStep draft={draft} onChange={updateDraft} />,
    <ReviewStep draft={draft} validationErrors={validationErrors} />,
  ][step];
  return (
    <>
      <div className="studio-page builder-page">
        <PageHeader
          eyebrow={
            configuration
              ? `Editing configuration · v${configuration.version}`
              : "New configuration · not saved yet"
          }
          title={draft.name}
          description=""
          actions={
            <>
              <span className="saved-state">
                <CheckCircle />
                {updateConfiguration.isPending || createConfiguration.isPending
                  ? "Saving…"
                  : data?.live
                    ? "Dataverse"
                    : "Local demo"}
              </span>
              <ToolbarButton
                type="button"
                appearance="subtle"
                icon={<Code />}
                onClick={() => setRecipeOpen(!recipeOpen)}
              >
                {recipeOpen ? "Hide recipe" : "View recipe"}
              </ToolbarButton>
              <ToolbarButton
                className="primary-command"
                appearance="primary"
                icon={<Play />}
                onClick={() => setTestRecordOpen(true)}
              >
                Test with a record
              </ToolbarButton>
            </>
          }
        />
        {recipeOpen && (
          <div className="recipe-drawer">
            <CompiledRecipe step={step} draft={draft} title="Compiled recipe" />
          </div>
        )}
        <div className="builder-layout">
          <aside
            className="step-rail bpf-process"
            aria-label="Configuration process"
          >
            <div className="bpf-process-header">
              <span>Configuration process</span>
              <em>
                Stage {step + 1} of {steps.length}
              </em>
            </div>
            <div className="bpf-stages">
              {steps.map((item, index) => {
                const Icon = item.icon;
                return (
                  <button
                    type="button"
                    key={item.title}
                    onClick={() => setStep(index)}
                    className={cn(
                      step === index && "active",
                      index < step && "complete",
                    )}
                    aria-current={step === index ? "step" : undefined}
                  >
                    <span>{index < step ? <Check /> : <Icon />}</span>
                    <div>
                      <small>
                        0{index + 1} · {item.stage}
                      </small>
                      <b>{item.title}</b>
                    </div>
                  </button>
                );
              })}
            </div>
          </aside>
          <main className="editor-panel">
            {content}
            <footer className="editor-footer" aria-label="Process actions">
              <Button
                variant="ghost"
                disabled={step === 0}
                onClick={() => setStep(step - 1)}
              >
                Previous
              </Button>
              <span>
                Step {step + 1} of {steps.length}
              </span>
              {step < 4 ? (
                <Button onClick={() => setStep(step + 1)}>
                  Continue <ArrowRight data-icon="inline-end" />
                </Button>
              ) : (
                <Button
                  onClick={publish}
                  disabled={
                    updateConfiguration.isPending ||
                    createConfiguration.isPending
                  }
                >
                  <Flow data-icon="inline-start" />
                  {updateConfiguration.isPending ||
                  createConfiguration.isPending
                    ? "Saving…"
                    : "Publish configuration"}
                </Button>
              )}
            </footer>
          </main>
        </div>
        {pane === "table" ? (
          <TableCatalogPane
            purpose="source"
            current={draft.entity}
            onClose={() => setPane(null)}
            onApply={(table) => {
              const sourceFields: string[] = [];
              const triggerColumns: string[] = [];
              const fetchXml = buildRecordSelectionFetchXml(
                table.logicalName,
                draft.maxRecords,
                table.primaryIdAttribute,
              );
              const relatedFetchXml = buildContextFetchXml(
                table.logicalName,
                sourceFields,
                [],
                table.primaryIdAttribute,
              );
              return save(
                {
                  ...draft,
                  entity: table.logicalName,
                  entitySetName: table.entitySetName,
                  entityIdField: table.primaryIdAttribute,
                  sourceFields,
                  triggerColumns,
                  relationships: [],
                  fetchXml,
                  relatedFetchXml,
                },
                {
                  csp_targetentity: table.logicalName,
                  csp_targetentityset: table.entitySetName,
                  csp_targetentityidfield: table.primaryIdAttribute,
                  csp_sourcefields: "[]",
                  csp_triggercolumns: "[]",
                  csp_relationships: "[]",
                  csp_fetchxml: fetchXml,
                  csp_relatedfetchxml: relatedFetchXml,
                },
              );
            }}
          />
        ) : pane === "fields" ? (
          <FieldsPane
            entity={draft.entity}
            entitySetName={
              draft.entitySetName ||
              getTargetEntityIdentity(draft.entity).entitySet
            }
            selectedFields={draft.sourceFields}
            onClose={() => setPane(null)}
            onApply={(fields) => updateDraft({ sourceFields: fields })}
          />
        ) : pane === "records" ? (
          <QueryDesignerPane
            purpose="records"
            targetEntity={draft.entity}
            entitySetName={
              draft.entitySetName ||
              getTargetEntityIdentity(draft.entity).entitySet
            }
            queryMode={draft.queryMode}
            initialViewId={
              draft.queryMode === 100000000
                ? draft.systemViewId
                : draft.userViewId
            }
            initialViewType={
              draft.queryMode === 100000001 ? "personal" : "system"
            }
            fetchXml={draft.fetchXml}
            maxRecords={draft.maxRecords}
            onClose={() => setPane(null)}
            onApply={({
              fetchXml,
              maxRecords,
              queryMode,
              viewId,
              viewType,
            }) => {
              const systemViewId = viewType === "system" ? viewId : undefined;
              const userViewId = viewType === "personal" ? viewId : undefined;
              return save(
                {
                  ...draft,
                  fetchXml,
                  maxRecords,
                  queryMode,
                  systemViewId,
                  userViewId,
                },
                {
                  csp_fetchxml: fetchXml,
                  csp_maxrecords: maxRecords,
                  csp_querymode: queryMode,
                  csp_systemviewid: systemViewId ?? null,
                  csp_userviewid: userViewId ?? null,
                },
              );
            }}
          />
        ) : pane === "relationships" ? (
          <QueryDesignerPane
            purpose="related"
            sourceEntity={draft.entity}
            sourceEntitySetName={
              draft.entitySetName ||
              getTargetEntityIdentity(draft.entity).entitySet
            }
            targetEntity={String(
              draft.relationships[0]?.entity ?? "activitypointer",
            )}
            entitySetName={String(
              draft.relationships[0]?.entitySetName ?? "activitypointers",
            )}
            relationship={draft.relationships[0]}
            queryMode={Number(
              draft.relationships[0]?.selectionMode ?? 100000002,
            )}
            initialViewId={
              draft.relationships[0]?.viewId
                ? String(draft.relationships[0].viewId)
                : undefined
            }
            initialViewType={
              draft.relationships[0]?.viewType === "personal"
                ? "personal"
                : "system"
            }
            fetchXml={relatedSelectionFetchXml(draft.relationships[0])}
            maxRecords={Number(draft.relationships[0]?.maxRecords ?? 12)}
            onClose={() => setPane(null)}
            onApply={({
              fetchXml,
              maxRecords,
              queryMode,
              viewId,
              viewType,
              relatedTable,
              relationship,
              fields,
            }) => {
              if (!relationship) return;
              const relationships = [
                {
                  ...relationship,
                  entitySetName: relatedTable?.entitySetName,
                  fields: fields ?? [],
                  maxRecords,
                  fetchXml,
                  selectionMode: queryMode,
                  ...(viewId ? { viewId, viewType } : {}),
                },
              ];
              const relatedFetchXml = buildContextFetchXml(
                draft.entity,
                draft.sourceFields,
                relationships,
                draft.entityIdField,
              );
              return save(
                { ...draft, relationships, relatedFetchXml },
                {
                  csp_relationships: JSON.stringify(relationships),
                  csp_relatedfetchxml: relatedFetchXml,
                },
              );
            }}
          />
        ) : pane === "contextQuery" ? (
          <ContextQueryPane
            fetchXml={draft.relatedFetchXml}
            onClose={() => setPane(null)}
            onApply={(relatedFetchXml) =>
              save(
                { ...draft, relatedFetchXml },
                { csp_relatedfetchxml: relatedFetchXml },
              )
            }
          />
        ) : pane === "outputTable" ? (
          <TableCatalogPane
            purpose="destination"
            current={draft.outputEntity}
            onClose={() => setPane(null)}
            onApply={(table) =>
              save(
                {
                  ...draft,
                  outputEntity: table.logicalName,
                  outputEntitySetName: table.entitySetName,
                  outputField: "",
                },
                { csp_outputentity: table.logicalName, csp_outputfield: "" },
              )
            }
          />
        ) : pane === "outputField" ? (
          <OutputFieldPane
            entity={draft.outputEntity}
            entitySetName={
              draft.outputEntitySetName ||
              getTargetEntityIdentity(draft.outputEntity).entitySet
            }
            current={draft.outputField}
            onClose={() => setPane(null)}
            onApply={(outputField) =>
              save({ ...draft, outputField }, { csp_outputfield: outputField })
            }
          />
        ) : (
          pane && (
            <EditPane
              pane={pane}
              current={paneCurrent}
              options={
                paneOptions.length ? paneOptions : paneContent[pane].options
              }
              onClose={() => setPane(null)}
              onApply={applyPane}
            />
          )
        )}
        {testRecordOpen && (
          <TestRecordPane
            draft={draft}
            onClose={() => setTestRecordOpen(false)}
          />
        )}
      </div>
    </>
  );
}

function Published() {
  const navigate = useNavigate();
  const { configurationId } = useParams();
  const { data } = useStudioRuntime();
  const configuration = data?.configurations.find(
    (item) => item.id === configurationId,
  );
  const builderPath = `/configurations/${configurationId ?? "new"}`;
  const flowCreated = Boolean(configuration?.flowId);
  const flowLabel =
    configuration?.flowDisplayName ||
    configuration?.flowName ||
    "Named by the generator flow after publishing";
  const flowUrl =
    flowCreated && data?.environmentId
      ? `https://make.powerautomate.com/environments/${data.environmentId}/flows/${configuration?.flowId}/details`
      : "";
  const promptLabel = configuration?.promptKey || configuration?.promptName || "Not selected";
  const modelLabel = configuration?.model || "AI Builder model";
  const flows = useCloudFlows();
  const notify = useStudioToast();
  const generatedFlows = configuration
    ? relatedCloudFlows(flows.data ?? [], configuration)
    : [];
  const activeEarlierVersions = generatedFlows.filter(
    (flow) => !flow.linked && flow.state === "Activated",
  ).length;
  const runNow = () => {
    if (!flowUrl) return;
    window.open(flowUrl, "_blank", "noopener");
    notify(
      "info",
      "Trigger the flow from Power Automate",
      "Use Run or edit the trigger there. New executions appear under Runs after a refresh.",
    );
  };
  return (
    <>
      <div className="studio-page publish-page">
        <div className="success-hero">
          <div className="success-mark">
            <Check />
          </div>
          <p className="micro-label text-brand-cyan">
            VERSIONED CONFIGURATION IN DATAVERSE
          </p>
          <h1>Recipe prepared</h1>
          <p>
            {flowCreated
              ? "The backend has generated the specialized Power Automate flow for this recipe."
              : "The configuration is stored and waiting for the backend to generate or update the specialized Power Automate flow."}
          </p>
          <div className="mt-6 flex justify-center gap-3">
            <Button
              onClick={() => navigate(builderPath, { state: { openTest: true } })}
            >
              <Play data-icon="inline-start" />
              Test configuration
            </Button>
            <Button
              variant="outline"
              onClick={() =>
                navigate(builderPath, { state: { openRecipe: true } })
              }
            >
              View JSON contract <ArrowRight data-icon="inline-end" />
            </Button>
            {flowUrl && (
              <Button variant="outline" onClick={runNow}>
                <Play data-icon="inline-start" />
                Run now
              </Button>
            )}
          </div>
          <div className="deployment-checks">
            <span>
              <CheckCircle />
              Schema validated
            </span>
            <span>
              <LinkIcon />
              AI Prompt bound
            </span>
            <span>
              <Blocks />
              Stored in the solution
            </span>
          </div>
        </div>
        <section className="publication-assets">
          <div className="publication-title">
            <div>
              <p className="micro-label text-muted-foreground">
                Compilation result
              </p>
              <h2>3 assets prepared</h2>
            </div>
            <span>
              <i />
              Demo · Irish Power Platform Summit 2026
            </span>
          </div>
          <div className="asset-timeline">
            <PublishedAsset
              icon={<Database />}
              title="Configuration stored"
              technical={`csp_aisummaryconfig · v${configuration?.version ?? "3.0"}`}
              detail="Versioned recipe available for audit."
              time="Ready"
            />
            <PublishedAsset
              icon={<Flow />}
              title="Flow definition prepared"
              technical={flowLabel}
              detail={
                flowCreated
                  ? "Cloud flow created and linked to this configuration."
                  : "Trigger, query, prompt, and write contract for the backend."
              }
              time={flowCreated ? "Created" : "Pending backend"}
              featured
            />
            <PublishedAsset
              icon={<Bot />}
              title="AI Prompt bound"
              technical={promptLabel}
              detail={`Inputs mapped · ${modelLabel}`}
              time={configuration?.promptId ? "Ready" : "Missing"}
            />
          </div>
        </section>
        <section className="publication-assets generated-flows" aria-label="Generated flows">
          <div className="publication-title">
            <div>
              <p className="micro-label text-muted-foreground">Power Automate</p>
              <h2>Generated flows</h2>
            </div>
            <span>
              <i />
              {flows.isLoading
                ? "Checking…"
                : `${generatedFlows.length} found`}
            </span>
          </div>
          {flows.isLoading ? (
            <p className="flows-empty">Checking the cloud flows in this environment…</p>
          ) : flows.isError ? (
            <p className="flows-empty">
              {errorMessage(flows.error, "Could not read the cloud flows.")}
            </p>
          ) : generatedFlows.length === 0 ? (
            <p className="flows-empty">
              No cloud flow has been generated for this recipe yet.
            </p>
          ) : (
            <ul className="flow-list">
              {generatedFlows.map((flow) => {
                const url = cloudFlowUrl(data?.environmentId ?? "", flow.id);
                return (
                  <li key={flow.id} className={flow.linked ? "linked" : ""}>
                    <span className="asset-symbol">
                      <Flow />
                    </span>
                    <div>
                      <b>{flow.name}</b>
                      <small>
                        {flow.state}
                        {flow.modifiedOn ? ` · updated ${formatRelativeDate(flow.modifiedOn)}` : ""}
                        {" · "}
                        <code>{flow.id}</code>
                      </small>
                    </div>
                    <Badge
                      variant="outline"
                      className={flow.linked ? "status-live" : "status-draft"}
                    >
                      {flow.linked ? "Linked" : "Earlier version"}
                    </Badge>
                    {url && (
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => window.open(url, "_blank", "noopener")}
                      >
                        Open <OpenIcon data-icon="inline-end" />
                      </Button>
                    )}
                  </li>
                );
              })}
            </ul>
          )}
          {activeEarlierVersions > 0 && (
            <p className="flows-warning">
              <Warning />
              <span>
                {activeEarlierVersions} earlier version
                {activeEarlierVersions === 1 ? " is" : "s are"} still active and may run in
                parallel with the linked flow. Turn them off in Power Automate or ask the
                backend to retire superseded flows when it regenerates.
              </span>
            </p>
          )}
        </section>
        <section className="execution-ready">
          <div>
            <p className="micro-label text-brand-blue">
              {flowCreated ? "Generated automation" : "Next step · backend"}
            </p>
            <h2>
              {flowCreated ? "Cloud flow is live" : "Contract ready to materialize"}
            </h2>
            <p>
              {flowCreated ? (
                <>
                  The generated flow <code>{flowLabel}</code> reads this recipe,
                  composes the context, runs the AI Prompt, and updates{" "}
                  <code>
                    {configuration?.outputEntity ?? "account"}.
                    {configuration?.outputField ?? "csp_aisummary"}
                  </code>
                  .
                </>
              ) : (
                <>
                  The backend reads the recipe, composes the context, and updates{" "}
                  <code>account.csp_aisummary</code>; the Code App already uses
                  the same Dataverse model.
                </>
              )}
            </p>
          </div>
          <div className="ready-flow">
            <span>
              <Database />
              <small>Recipe</small>
            </span>
            <i />
            <span>
              <Flow />
              <small>Backend</small>
            </span>
            <i />
            <span>
              <Bot />
              <small>AI Prompt</small>
            </span>
            <i />
            <span>
              <Target />
              <small>Summary</small>
            </span>
          </div>
          <Button variant="outline" onClick={() => navigate("/")}>
            Back to configurations
          </Button>
        </section>
      </div>
    </>
  );
}

function PublishedAsset({
  icon,
  title,
  technical,
  detail,
  time,
  featured,
}: {
  icon: ReactNode;
  title: string;
  technical: string;
  detail: string;
  time: string;
  featured?: boolean;
}) {
  return (
    <div className={cn("published-asset", featured && "featured")}>
      <span className="asset-check">
        <Check />
      </span>
      <div className="asset-symbol">{icon}</div>
      <div>
        <p>{title}</p>
        <code>{technical}</code>
        <small>{detail}</small>
      </div>
      <em>{time}</em>
    </div>
  );
}

type PromptDraft = StudioPromptInput;

const promptModelOptions = ["gpt-4.1-mini", "gpt-4.1", "gpt-4o-mini", "gpt-4o"];

const newPromptTemplate = `You are preparing a concise account summary for account managers.

Objective
Summarize the current situation of the account and highlight what needs attention before the next customer conversation.

Source context
Use {{account_context}} as the only source of truth. Never invent data; if something is missing, say so explicitly.

Output
- Situation: two or three sentences.
- Risks and opportunities: short bullet list grounded in the context.
- Recommended next actions: up to three bullets.

Write in English with a direct, professional tone.`;

function promptDraftFrom(prompt?: StudioPrompt | null): PromptDraft {
  return {
    name: prompt?.name ?? "",
    key: prompt?.key ?? "",
    model: prompt?.model || "gpt-4.1-mini",
    version: prompt?.version || "1.0",
    content: prompt ? prompt.content : newPromptTemplate,
  };
}

function slugifyPromptKey(value: string) {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);
}

function PromptEditorPane({
  prompt,
  usedBy,
  saving,
  error,
  onClose,
  onSave,
}: {
  prompt: StudioPrompt | null;
  usedBy: StudioConfiguration[];
  saving: boolean;
  error: string;
  onClose: () => void;
  onSave: (draft: PromptDraft) => void;
}) {
  const [draft, setDraft] = useState<PromptDraft>(() => promptDraftFrom(prompt));
  const [keyTouched, setKeyTouched] = useState(Boolean(prompt));
  const update = <K extends keyof PromptDraft>(key: K, value: PromptDraft[K]) =>
    setDraft((current) => ({ ...current, [key]: value }));
  const setName = (value: string) =>
    setDraft((current) => ({
      ...current,
      name: value,
      key: keyTouched ? current.key : slugifyPromptKey(value),
    }));
  const valid =
    draft.name.trim().length > 0 &&
    draft.key.trim().length > 0 &&
    draft.content.trim().length > 0;
  const estimatedTokens = estimateRecipeTokens({
    promptContent: draft.content,
    sourceFields: [],
    relationships: [],
  });
  return (
    <>
      <button
        className="pane-scrim"
        type="button"
        aria-label="Close edit pane"
        onClick={onClose}
      />
      <aside
        className="edit-pane prompt-editor-pane"
        aria-label={prompt ? "Edit AI Prompt" : "New AI Prompt"}
      >
        <header>
          <div>
            <p className="micro-label text-brand-blue">AI Prompt library</p>
            <h2>{prompt ? "Edit prompt" : "New prompt"}</h2>
          </div>
          <Button type="button" variant="ghost" size="icon" aria-label="Close pane" onClick={onClose}>
            <Dismiss />
          </Button>
        </header>
        {prompt && (
          <div className="pane-current">
            <span>
              <Bot />
            </span>
            <div>
              <small>Stored in csp_aiprompt</small>
              <b>{prompt.name}</b>
              <code>{prompt.id}</code>
            </div>
          </div>
        )}
        <section>
          <label className="studio-field">
            <span>Name</span>
            <Input
              appearance="outline"
              aria-label="Prompt name"
              value={draft.name}
              onChange={(event) => setName(event.target.value)}
              placeholder="Account executive summary"
            />
            {!prompt && (
              <small>Example: "Account executive summary" or "Case escalation brief".</small>
            )}
          </label>
          <div className="grid grid-cols-2 gap-3">
            <label className="studio-field">
              <span>Key</span>
              <Input
                appearance="outline"
                aria-label="Prompt key"
                value={draft.key}
                onChange={(event) => {
                  setKeyTouched(true);
                  update("key", event.target.value);
                }}
                placeholder="account-operations"
              />
              <small>Stable identifier used by generated flows.</small>
            </label>
            <label className="studio-field">
              <span>Version</span>
              <Input
                appearance="outline"
                aria-label="Prompt version"
                value={draft.version}
                onChange={(event) => update("version", event.target.value)}
              />
            </label>
          </div>
          <label className="studio-field">
            <span>Model</span>
            <Select
              aria-label="Prompt model"
              value={
                promptModelOptions.includes(draft.model) ? draft.model : "custom"
              }
              onChange={(event) => {
                if (event.target.value !== "custom") update("model", event.target.value);
              }}
            >
              {promptModelOptions.map((model) => (
                <option key={model} value={model}>
                  {model}
                </option>
              ))}
              {!promptModelOptions.includes(draft.model) && (
                <option value="custom">{draft.model}</option>
              )}
            </Select>
          </label>
          <label className="studio-field">
            <span>Instructions</span>
            <Textarea
              resize="vertical"
              className="prompt-library-content"
              aria-label="Prompt instructions content"
              value={draft.content}
              onChange={(event) => update("content", event.target.value)}
            />
            <small>
              Use <code>{"{{account_context}}"}</code> for the compiled Dataverse
              context · ~{estimatedTokens.toLocaleString("en-US")} tokens
            </small>
          </label>
          {prompt && (
            <div className="prompt-usage">
              <p className="field-caption">Used by</p>
              {usedBy.length === 0 ? (
                <small>No configuration is bound to this prompt yet.</small>
              ) : (
                <ul>
                  {usedBy.map((configuration) => (
                    <li key={configuration.id}>
                      <Link to={`/configurations/${configuration.id}`}>
                        {configuration.name}
                      </Link>
                      <small>
                        {configuration.status === 100000001 ? "Published" : "Draft"}
                      </small>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          )}
          {error && (
            <p className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
              {error}
            </p>
          )}
        </section>
        <footer>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button disabled={!valid || saving} onClick={() => onSave(draft)}>
            {saving ? "Saving…" : prompt ? "Save changes" : "Create prompt"}
          </Button>
        </footer>
      </aside>
    </>
  );
}

function Prompts() {
  const { data, refetch, isFetching } = useStudioRuntime();
  const prompts = data?.prompts ?? [];
  const configurations = data?.configurations ?? [];
  const [editing, setEditing] = useState<StudioPrompt | "new" | null>(null);
  const [error, setError] = useState("");
  const savePrompt = useSaveStudioPrompt();
  const createPrompt = useCreateStudioPrompt();
  const usedBy = (promptId: string) =>
    configurations.filter(
      (configuration) =>
        configuration.promptId.toLowerCase() === promptId.toLowerCase(),
    );
  const notify = useStudioToast();
  const handleSave = async (draft: PromptDraft) => {
    setError("");
    try {
      if (editing === "new") {
        await createPrompt.mutateAsync(draft);
        notify("success", "Prompt created", `"${draft.name}" is available to every configuration.`);
      } else if (editing) {
        await savePrompt.mutateAsync({ id: editing.id, changes: draft });
        notify("success", "Prompt saved", `"${draft.name}" was updated in Dataverse.`);
      }
      setEditing(null);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The prompt could not be saved.");
    }
  };
  return (
    <>
      <div className="studio-page catalog-page">
        <PageHeader
          eyebrow="Summary Studio"
          title="AI Prompts"
          description="Reusable prompt instructions stored in Dataverse. Editing a prompt updates every configuration bound to it."
        />
        <Toolbar aria-label="Prompt commands" className="catalog-command-bar">
          <ToolbarButton
            className="primary-command"
            appearance="primary"
            icon={<Sparkles />}
            onClick={() => {
              setError("");
              setEditing("new");
            }}
          >
            New prompt
          </ToolbarButton>
          <ToolbarButton
            appearance="subtle"
            icon={<RefreshCw />}
            onClick={() => refetch()}
            disabled={isFetching}
          >
            {isFetching ? "Refreshing" : "Refresh"}
          </ToolbarButton>
          <ToolbarDivider />
          <span className="catalog-command-hint">
            Stored in <code>csp_aiprompt</code>
          </span>
        </Toolbar>
        <section className="catalog-grid" aria-labelledby="prompt-grid-title">
          <div className="catalog-grid-heading">
            <div>
              <h2 id="prompt-grid-title">Prompt library</h2>
              <p>Select a prompt to review or edit its instructions.</p>
            </div>
            <Badge variant="outline">{prompts.length} items</Badge>
          </div>
          <Table aria-label="AI Prompts" className="config-table" size="medium">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Prompt</TableHeaderCell>
                <TableHeaderCell>Model</TableHeaderCell>
                <TableHeaderCell>Instructions</TableHeaderCell>
                <TableHeaderCell>Used by</TableHeaderCell>
                <TableHeaderCell aria-label="Open" />
              </TableRow>
            </TableHeader>
            <TableBody>
              {prompts.map((prompt) => {
                const bound = usedBy(prompt.id);
                return (
                  <TableRow key={prompt.id}>
                    <TableCell>
                      <TableCellLayout media={<span className="entity-icon"><Bot /></span>}>
                        <button
                          type="button"
                          className="configuration-link text-left"
                          onClick={() => {
                            setError("");
                            setEditing(prompt);
                          }}
                        >
                          <b>{prompt.name}</b>
                          <small>{prompt.key || "no key"} · v{prompt.version || "1.0"}</small>
                        </button>
                      </TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <span className="table-value">
                        <b>{prompt.model || "Not set"}</b>
                        <small>AI Builder model</small>
                      </span>
                    </TableCell>
                    <TableCell>
                      <span className="table-value">
                        <b className="prompt-preview">{prompt.content || "No instructions yet"}</b>
                        <small>
                          {prompt.content.length.toLocaleString("en-US")} characters
                        </small>
                      </span>
                    </TableCell>
                    <TableCell>
                      <span className="table-value">
                        <b>
                          {bound.length} configuration{bound.length === 1 ? "" : "s"}
                        </b>
                        <small>{bound.map((item) => item.name).join(", ") || "Unused"}</small>
                      </span>
                    </TableCell>
                    <TableCell>
                      <Button
                        variant="ghost"
                        size="icon-sm"
                        aria-label={`Open ${prompt.name}`}
                        onClick={() => {
                          setError("");
                          setEditing(prompt);
                        }}
                      >
                        <ChevronRight />
                      </Button>
                    </TableCell>
                  </TableRow>
                );
              })}
              {prompts.length === 0 && (
                <TableRow>
                  <TableCell colSpan={5}>
                    <div className="catalog-empty-state">
                      <Bot />
                      <span>
                        <b>No AI Prompts yet.</b>
                        <small>Create one to reuse it across configurations.</small>
                      </span>
                    </div>
                  </TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>
        </section>
      </div>
      {editing && (
        <PromptEditorPane
          key={editing === "new" ? "new" : editing.id}
          prompt={editing === "new" ? null : editing}
          usedBy={editing === "new" ? [] : usedBy(editing.id)}
          saving={savePrompt.isPending || createPrompt.isPending}
          error={error}
          onClose={() => setEditing(null)}
          onSave={handleSave}
        />
      )}
    </>
  );
}

function Runs() {
  const { data, refetch, isFetching } = useStudioRuntime();
  const allRuns = data?.runs ?? [];
  const configurations = data?.configurations ?? [];
  const [configurationFilter, setConfigurationFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [expandedRun, setExpandedRun] = useState<string | null>(null);
  const statusOf = (run: StudioRun) =>
    run.status === 100000000
      ? "Successful"
      : run.status === 100000001
        ? "Running"
        : "Error";
  const runs = allRuns.filter(
    (run) =>
      (configurationFilter === "all" || run.configurationId === configurationFilter) &&
      (statusFilter === "all" || statusOf(run) === statusFilter),
  );
  const successes = runs.filter((run) => run.status === 100000000).length;
  const failures = runs.filter((run) => statusOf(run) === "Error").length;
  const average = runs.length
    ? runs.reduce((total, run) => total + run.latencyMs, 0) / runs.length
    : 0;
  const tokens = runs.reduce((total, run) => total + run.tokens, 0);
  const configurationById = new Map(configurations.map((item) => [item.id, item]));
  const formatDate = (value: string | null) =>
    value
      ? new Intl.DateTimeFormat("en-IE", {
          day: "2-digit",
          month: "short",
          hour: "2-digit",
          minute: "2-digit",
        }).format(new Date(value))
      : "Not available";
  return (
    <>
      <div className="studio-page">
        <PageHeader
          eyebrow="Operations"
          title="Runs"
          description="Inspect which recipe ran, which record it processed, and what result was written to Dataverse."
          actions={
            <ToolbarButton
              appearance="subtle"
              icon={<RefreshCw />}
              onClick={() => refetch()}
              disabled={isFetching}
            >
              {isFetching ? "Refreshing" : "Refresh"}
            </ToolbarButton>
          }
        />
        <div className="runs-filters" role="group" aria-label="Run filters">
          <label>
            <span>Configuration</span>
            <Select
              aria-label="Filter by configuration"
              value={configurationFilter}
              onChange={(event) => setConfigurationFilter(event.target.value)}
            >
              <option value="all">All configurations</option>
              {configurations.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.name}
                </option>
              ))}
            </Select>
          </label>
          <label>
            <span>Status</span>
            <Select
              aria-label="Filter by status"
              value={statusFilter}
              onChange={(event) => setStatusFilter(event.target.value)}
            >
              <option value="all">All statuses</option>
              <option value="Successful">Successful</option>
              <option value="Running">Running</option>
              <option value="Error">Error</option>
            </Select>
          </label>
          <small>
            {runs.length} of {allRuns.length} runs
          </small>
        </div>
        <div className="run-summary">
          <div>
            <ActivityStat value={String(runs.length)} label="Recorded" />
            <ActivityStat
              value={`${runs.length ? ((successes / runs.length) * 100).toFixed(1) : "0"}%`}
              label="Successful"
            />
            <ActivityStat value={String(failures)} label="Errors" />
            <ActivityStat
              value={`${(average / 1000).toFixed(1)} s`}
              label="Average duration"
            />
            <ActivityStat
              value={tokens.toLocaleString("en-IE")}
              label="Tokens"
            />
          </div>
          <span>
            <i />
            {data?.live ? "Dataverse connected" : "Local demo"}
          </span>
        </div>
        <div className="runs-table-surface">
          <Table aria-label="Summary runs" className="runs-table" size="medium">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Run</TableHeaderCell>
                <TableHeaderCell>Configuration</TableHeaderCell>
                <TableHeaderCell>Record</TableHeaderCell>
                <TableHeaderCell>Status</TableHeaderCell>
                <TableHeaderCell>Duration</TableHeaderCell>
                <TableHeaderCell>Started</TableHeaderCell>
                <TableHeaderCell aria-label="Flow run" />
              </TableRow>
            </TableHeader>
            <TableBody>
              {runs.map((run) => {
                const status = statusOf(run);
                const configuration = configurationById.get(run.configurationId);
                const runUrl = cloudFlowRunUrl(
                  data?.environmentId ?? "",
                  configuration?.flowId ?? "",
                  run.flowRunId,
                );
                const expanded = expandedRun === run.id;
                return (
                  <Fragment key={run.id}>
                    <TableRow
                      className={status === "Error" ? "run-row-error" : undefined}
                    >
                      <TableCell><code>{run.name}</code></TableCell>
                      <TableCell>
                        <span className="table-value">
                          <b>{run.configurationName}</b>
                          <small>{run.promptName || run.model || "AI Prompt"}</small>
                        </span>
                      </TableCell>
                      <TableCell>
                        <span className="table-value">
                          <b>{run.accountName}</b>
                          <small>{run.targetEntity || configuration?.entity || "record"}</small>
                        </span>
                      </TableCell>
                      <TableCell>
                        <span className="table-value">
                          <Badge
                            variant="outline"
                            className={
                              status === "Successful"
                                ? "status-live"
                                : status === "Error"
                                  ? "status-error"
                                  : "status-draft"
                            }
                          >
                            {status}
                          </Badge>
                          {run.error && (
                            <button
                              type="button"
                              className="run-error-toggle"
                              aria-expanded={expanded}
                              onClick={() => setExpandedRun(expanded ? null : run.id)}
                            >
                              <ErrorCircle />
                              {expanded ? "Hide error" : "Show error"}
                            </button>
                          )}
                        </span>
                      </TableCell>
                      <TableCell>{(run.latencyMs / 1000).toFixed(1)} s</TableCell>
                      <TableCell>{formatDate(run.timestamp)}</TableCell>
                      <TableCell>
                        {runUrl ? (
                          <Button
                            variant="ghost"
                            size="icon-sm"
                            aria-label={`Open flow run ${run.name} in Power Automate`}
                            title="Open flow run in Power Automate"
                            onClick={() => window.open(runUrl, "_blank", "noopener")}
                          >
                            <OpenIcon />
                          </Button>
                        ) : null}
                      </TableCell>
                    </TableRow>
                    {expanded && run.error && (
                      <TableRow className="run-error-row">
                        <TableCell colSpan={7}>
                          <pre className="run-error">{run.error}</pre>
                        </TableCell>
                      </TableRow>
                    )}
                  </Fragment>
                );
              })}
              {runs.length === 0 && (
                <TableRow>
                  <TableCell colSpan={7}>
                    <div className="catalog-empty-state">
                      <History />
                      <span>
                        <b>No runs match these filters.</b>
                        <small>Runs are recorded in csp_aiusage by the generated flows.</small>
                      </span>
                    </div>
                  </TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>
        </div>
      </div>
    </>
  );
}

function Summaries() {
  const { data, refetch, isFetching } = useStudioRuntime();
  const configurations = data?.configurations ?? [];
  const [selectedId, setSelectedId] = useState("");
  const configuration =
    configurations.find((item) => item.id === selectedId) ?? configurations[0];
  const outputTable = useMemo<DataverseTableMetadata | undefined>(
    () =>
      configuration
        ? {
            logicalName: configuration.outputEntity,
            entitySetName:
              configuration.outputEntitySetName ||
              getTargetEntityIdentity(configuration.outputEntity).entitySet,
            primaryIdAttribute:
              configuration.outputEntity === configuration.entity
                ? configuration.entityIdField
                : getTargetEntityIdentity(configuration.outputEntity).idField,
            displayName: configuration.outputEntity,
          }
        : undefined,
    [configuration],
  );
  const outputMetadata = useDataverseTableMetadata(outputTable);
  const nameAttribute = outputMetadata.data?.primaryNameAttribute ?? "name";
  const summaries = useOutputSummaries(configuration, nameAttribute, data?.accounts ?? []);
  const [expanded, setExpanded] = useState<string | null>(null);
  const items = summaries.data ?? [];
  const latest = items
    .map((item) => item.generatedOn)
    .filter((value): value is string => Boolean(value))
    .sort()
    .at(-1);
  const formatDate = (value: string | null) =>
    value
      ? new Intl.DateTimeFormat("en-IE", {
          day: "2-digit",
          month: "short",
          hour: "2-digit",
          minute: "2-digit",
        }).format(new Date(value))
      : "Unknown";
  return (
    <>
      <div className="studio-page catalog-page">
        <PageHeader
          eyebrow="Results"
          title="Generated summaries"
          description="Records whose destination column already holds an AI summary written by the generated flow."
          actions={
            <ToolbarButton
              appearance="subtle"
              icon={<RefreshCw />}
              onClick={() => {
                refetch();
                summaries.refetch();
              }}
              disabled={isFetching || summaries.isFetching}
            >
              {isFetching || summaries.isFetching ? "Refreshing" : "Refresh"}
            </ToolbarButton>
          }
        />
        <div className="runs-filters" role="group" aria-label="Summary filters">
          <label>
            <span>Configuration</span>
            <Select
              aria-label="Configuration"
              value={configuration?.id ?? ""}
              onChange={(event) => setSelectedId(event.target.value)}
            >
              {configurations.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.name}
                </option>
              ))}
            </Select>
          </label>
          {configuration && (
            <small>
              Reading <code>{configuration.outputEntity}.{configuration.outputField}</code>
              {" · "}
              {items.length} record{items.length === 1 ? "" : "s"} with a summary
              {latest ? ` · latest ${formatRelativeDate(latest)}` : ""}
            </small>
          )}
        </div>
        {summaries.isLoading ? (
          <p className="flows-empty">Loading generated summaries…</p>
        ) : summaries.isError ? (
          <p className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
            {errorMessage(summaries.error, "Could not load the generated summaries.")}
          </p>
        ) : items.length === 0 ? (
          <div className="catalog-grid">
            <div className="catalog-empty-state">
              <DocumentText />
              <span>
                <b>No summaries yet.</b>
                <small>Run the generated flow and refresh to see the results here.</small>
              </span>
            </div>
          </div>
        ) : (
          <div className="summary-list">
            {items.map((item) => {
              const open = expanded === item.id;
              return (
                <article
                  key={item.id}
                  className={cn("summary-card", open && "open")}
                >
                  <header>
                    <span className="entity-icon">
                      <DocumentText />
                    </span>
                    <div>
                      <b>{item.name}</b>
                      <small>
                        Generated {formatDate(item.generatedOn)} · <code>{item.id}</code>
                      </small>
                    </div>
                    <Button
                      variant="ghost"
                      size="sm"
                      aria-expanded={open}
                      aria-label={`${open ? "Collapse" : "Expand"} summary for ${item.name}`}
                      onClick={() => setExpanded(open ? null : item.id)}
                    >
                      {open ? "Collapse" : "Expand"}
                    </Button>
                  </header>
                  <p className="summary-text">{item.summary}</p>
                </article>
              );
            })}
          </div>
        )}
      </div>
    </>
  );
}
function ActivityStat({ value, label }: { value: string; label: string }) {
  return (
    <span>
      <b>{value}</b>
      <small>{label}</small>
    </span>
  );
}

export default function App() {
  return (
    <Shell>
      <Routes>
        <Route path="/" element={<Overview />} />
        <Route path="/configurations/new" element={<Builder />} />
        <Route path="/configurations/:configurationId" element={<Builder />} />
        <Route path="/published/:configurationId" element={<Published />} />
        <Route path="/prompts" element={<Prompts />} />
        <Route path="/summaries" element={<Summaries />} />
        <Route path="/runs" element={<Runs />} />
      </Routes>
    </Shell>
  );
}
