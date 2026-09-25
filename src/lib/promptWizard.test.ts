import { describe, expect, it } from "vitest";
import { buildPromptDraft, type PromptWizardAnswers } from "./promptWizard";

const answers: PromptWizardAnswers = {
  objective: "Give account managers a concise operational briefing before a customer call.",
  audience: "Account managers",
  format: "Structured sections",
  length: "Concise",
  tone: "Direct and professional",
  language: "English",
  includeActions: true,
  citeEvidence: true,
  handleMissingData: true,
};

describe("prompt design wizard", () => {
  it("turns business choices and Dataverse context into an editable prompt", () => {
    const prompt = buildPromptDraft({
      answers,
      contextVariable: "{{account_context}}",
      entityLabel: "Account",
      fields: ["name", "revenue", "description"],
      relatedSources: ["activitypointer"],
    });

    expect(prompt).toContain(answers.objective);
    expect(prompt).toContain("Audience: Account managers");
    expect(prompt).toContain("{{account_context}}");
    expect(prompt).toContain("name, revenue, description");
    expect(prompt).toContain("Recommended next actions");
    expect(prompt).toContain("Do not invent facts");
  });
});
