import { fireEvent, render, screen } from "@testing-library/react";
import { MemoryRouter } from "react-router-dom";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { describe, expect, it } from "vitest";
import App from "./App";

const renderRoute = (route: string) => {
  const queryClient = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  return render(
    <QueryClientProvider client={queryClient}>
      <MemoryRouter initialEntries={[route]}>
        <App />
      </MemoryRouter>
    </QueryClientProvider>,
  );
};

describe("Creativity Spark Summary Studio", () => {
  it("positions the product as a configurable summary automation studio", async () => {
    renderRoute("/");

    expect(await screen.findByText("Summary Studio")).toBeInTheDocument();
    expect(
      screen.getByRole("heading", { name: "Configure AI summaries" }),
    ).toBeInTheDocument();
    expect(await screen.findByText("Account operations summary")).toBeInTheDocument();
    expect(screen.getByLabelText("Studio summary")).toHaveTextContent(/1\s*published configurations/);
    expect(screen.getByRole("link", { name: /New configuration/ })).toBeInTheDocument();
    expect(screen.getByText("Foundry evaluations")).toBeInTheDocument();
  });

  it("opens each configuration by its Dataverse identifier", async () => {
    renderRoute("/");

    const configurationLink = await screen.findByRole("link", { name: /Account operations summary/ });
    expect(configurationLink).toHaveAttribute("href", "/configurations/mock-config");
  });

  it("starts a clean draft on the new configuration route", async () => {
    renderRoute("/configurations/new");

    expect(await screen.findByRole("heading", { name: "New summary" })).toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Account operations summary" })).not.toBeInTheDocument();
  });

  it("exposes the five decisions required to define a summary", async () => {
    renderRoute("/configurations/account-operations");

    expect(await screen.findByRole("heading", { name: "Account operations summary" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Data and context/ })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Prompt and model/ })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Destination/ })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Execution/ })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Review and publish/ })).toBeInTheDocument();
    expect(screen.getByText("3 demo records")).toBeInTheDocument();
  });

  it("guides configuration through a familiar Dataverse business process flow", async () => {
    renderRoute("/configurations/account-operations");

    expect(await screen.findByLabelText("Configuration process")).toBeInTheDocument();
    expect(screen.getByText("01 · Context")).toBeInTheDocument();
    expect(screen.getByText("02 · Prompt")).toBeInTheDocument();
    expect(screen.getByText("03 · Destination")).toBeInTheDocument();
    expect(screen.getByText("04 · Automation")).toBeInTheDocument();
    expect(screen.getByText("05 · Publish")).toBeInTheDocument();
  });

  it("makes every data and context setting explicitly configurable", async () => {
    renderRoute("/configurations/account-operations");

    expect(await screen.findByRole("region", { name: "1. Records to summarize" })).toBeInTheDocument();
    expect(screen.getByRole("region", { name: "2. Context for each record" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Change source table" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Configure record FetchXML" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Change summary mode" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Configure source fields" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Configure context FetchXML" })).toBeInTheDocument();
    expect(screen.getByText("csp_fetchxml")).toBeInTheDocument();
    expect(screen.getAllByText("csp_relatedfetchxml").length).toBeGreaterThan(0);
  });

  it("opens the context FetchXML editor from the context section", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: "Configure context FetchXML" }));
    expect(screen.getByRole("complementary", { name: "Edit context query" })).toBeInTheDocument();
    expect((screen.getByLabelText("Context FetchXML") as HTMLTextAreaElement).value).toContain("{{recordId}}");
  });

  it("offers filters, saved views, and FetchXML for record selection", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: "Configure record FetchXML" }));
    expect(screen.getByRole("complementary", { name: "Configure record selection" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Build filters" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Choose a view" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Edit FetchXML" })).toBeInTheDocument();
  });

  it("offers the same query modes for related records", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: "Configure related records" }));
    expect(screen.getByRole("complementary", { name: "Configure related records" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Build filters" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Choose a view" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Edit FetchXML" })).toBeInTheDocument();
  });

  it("opens a model-driven side pane when a configuration property is edited", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: "Change source table" }));

    expect(screen.getByRole("complementary", { name: "Edit primary table" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Primary table" })).toBeInTheDocument();
    expect(screen.getAllByText("Account · account")).toHaveLength(2);
  });

  it("adds a Dataverse source field from the side pane", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: "Configure source fields" }));
    expect(screen.getByRole("complementary", { name: "Edit selected fields" })).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: /Account number/ }));
    fireEvent.click(screen.getByRole("button", { name: "Apply changes" }));

    expect(screen.getByText("accountnumber")).toBeInTheDocument();
  });

  it("edits the destination through model-driven side panes", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Destination/ }));
    fireEvent.click(screen.getByRole("button", { name: "Edit destination column" }));
    expect(screen.getByRole("complementary", { name: "Edit destination column" })).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: /Description/ }));
    fireEvent.click(screen.getByRole("button", { name: "Apply changes" }));
    expect(screen.getByText("description")).toBeInTheDocument();
  });

  it("changes the execution pattern from the execution stage", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Execution/ }));
    fireEvent.click(screen.getByRole("button", { name: "Scheduled" }));

    expect(screen.getByRole("button", { name: "Scheduled" })).toHaveAttribute("aria-pressed", "true");
    expect(screen.getByText("Daily schedule")).toBeInTheDocument();
  });

  it("keeps the publish action in the process footer", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Review and publish/ }));
    const footer = screen.getByRole("contentinfo", { name: "Process actions" });
    expect(footer).toContainElement(screen.getByRole("button", { name: "Publish configuration" }));
  });

  it("keeps technical recipe detail on demand so the builder remains focused", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: "View recipe" }));

    expect(screen.getByRole("heading", { name: "Compiled recipe" })).toBeInTheDocument();
    expect(screen.getByText("account.csp_aisummary")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Hide recipe" })).toBeInTheDocument();
  });

  it("shows a maker-friendly prompt editor with explicit runtime inputs", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Prompt and model/ }));
    expect(screen.getByText("Account executive summary")).toBeInTheDocument();
    expect(screen.getByText("GPT-4.1 mini")).toBeInTheDocument();
    expect(screen.getAllByText("{{account_context}}")).toHaveLength(2);
    expect(screen.getByText("Estimate · 1,240 tokens per run")).toBeInTheDocument();
  });

  it("helps the maker refine the prompt from the selected Dataverse context", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Prompt and model/ }));
    fireEvent.click(screen.getByRole("button", { name: /Help me improve this prompt/ }));

    expect(screen.getByText("Prompt assistant")).toBeInTheDocument();
    expect(screen.getByText("Context detected · Account + activities")).toBeInTheDocument();
    expect(screen.getByText("Reduce non-essential context")).toBeInTheDocument();
    expect(screen.getByText("−18% estimated tokens")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Apply refined version" })).toBeInTheDocument();
  });

  it("previews the specialized cloud flow contract generated from the recipe", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Execution/ }));
    expect(screen.getAllByText("When an account changes")).toHaveLength(2);
    expect(screen.getByText("Flow contract for the backend")).toBeInTheDocument();
    expect(screen.getByText("Query Dataverse")).toBeInTheDocument();
    expect(screen.getByText("Run AI Prompt")).toBeInTheDocument();
    expect(screen.getByText("Save summary")).toBeInTheDocument();
  });

  it("publishes the configuration and presents the prepared Power Platform assets", async () => {
    renderRoute("/configurations/account-operations");

    fireEvent.click(await screen.findByRole("button", { name: /Review and publish/ }));
    fireEvent.click(screen.getByRole("button", { name: "Publish configuration" }));

    expect(await screen.findByRole("heading", { name: "Recipe prepared" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /View JSON contract/i })).toHaveClass("text-white");
    expect(screen.getByText("csp_SUM_AccountOperations_v3")).toBeInTheDocument();
    expect(screen.getByText("Configuration stored")).toBeInTheDocument();
    expect(screen.getByText("Flow definition prepared")).toBeInTheDocument();
    expect(screen.getAllByText("AI Prompt bound").length).toBeGreaterThan(0);
    expect(screen.getByText("Schema validated")).toBeInTheDocument();
    expect(screen.getByText("Stored in the solution")).toBeInTheDocument();
  });
});
