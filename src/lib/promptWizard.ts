export interface PromptWizardAnswers {
  objective: string;
  audience: string;
  format: string;
  length: string;
  tone: string;
  language: string;
  includeActions: boolean;
  citeEvidence: boolean;
  handleMissingData: boolean;
}

interface PromptDraftInput {
  answers: PromptWizardAnswers;
  contextVariable: string;
  entityLabel: string;
  fields: string[];
  relatedSources: string[];
}

export function buildPromptDraft({ answers, contextVariable, entityLabel, fields, relatedSources }: PromptDraftInput) {
  const outputSections = [
    "Current situation",
    "Business impact and risks",
    answers.includeActions ? "Recommended next actions" : null,
  ].filter(Boolean).join("\n- ");
  const evidenceRule = answers.citeEvidence
    ? "Connect every important statement to a specific value or event in the supplied context."
    : "Use only information that is present in the supplied context.";
  const missingDataRule = answers.handleMissingData
    ? "Do not invent facts. If information is missing or uncertain, say so explicitly."
    : "Do not add facts that are not present in the context.";
  const relatedContext = relatedSources.length ? relatedSources.join(", ") : "none";

  return `You are preparing a ${answers.length.toLowerCase()} ${entityLabel.toLowerCase()} summary.

Objective
${answers.objective.trim()}

Audience: ${answers.audience}
Language: ${answers.language}
Tone: ${answers.tone}
Output format: ${answers.format}

Source context
Use ${contextVariable} as the only source of truth.
Available fields: ${fields.join(", ")}.
Related Dataverse records: ${relatedContext}.

Required output
- ${outputSections}

Rules
- ${evidenceRule}
- ${missingDataRule}
- Prefer concrete facts, dates, amounts, statuses, and named owners over generic statements.
- Remove repetition and omit sections that have no useful information.
- Return only the finished summary; do not describe your reasoning.`;
}
