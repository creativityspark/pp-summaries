/**
 * FetchXML construction helpers for the Visual Filter Builder.
 *
 * The builder state is intentionally narrow:
 *   - one entity
 *   - one top-level AND/OR filter group
 *   - 0-N conditions inside that group
 *   - 0-1 sort
 *   - one top-N
 *
 * Nested groups, multi-entity joins, link-entity, and aggregates are NOT
 * supported by the visual builder — use the Paste FetchXML tab for those.
 */

export type FilterMode = "and" | "or";

export interface FilterCondition {
  /** Logical name of the attribute, e.g. "modifiedon", "name". */
  attribute: string;
  /** FetchXML operator name, e.g. "eq", "last-x-days", "null". */
  operator: string;
  /** Raw value for operators that take one. Empty for null/not-null/today/etc. */
  value: string;
}

export interface FetchXmlBuilderState {
  entity: string;
  filterMode: FilterMode;
  conditions: FilterCondition[];
  sortAttribute: string;
  sortDescending: boolean;
  top: number;
}

/**
 * Operators that do not need a value at all.
 */
export const NO_VALUE_OPERATORS = new Set([
  "null",
  "not-null",
  "today",
  "yesterday",
  "tomorrow",
  "this-week",
  "this-month",
  "this-year",
  "last-week",
  "last-month",
  "last-year",
  "next-week",
  "next-month",
  "next-year",
  "eq-userid",
  "ne-userid",
]);

/**
 * High-level attribute type buckets the UI uses to choose pickers and
 * filter operator lists. Maps from the Dataverse AttributeType enum.
 */
export type AttributeBucket =
  | "string"
  | "memo"
  | "integer"
  | "decimal"
  | "money"
  | "datetime"
  | "boolean"
  | "picklist"
  | "lookup"
  | "unknown";

/**
 * Operators per attribute type. Curated to the ones that are useful in
 * the demo and don't require multi-value or range UI (`in`, `between`...).
 */
export const OPERATORS_BY_BUCKET: Record<
  AttributeBucket,
  { value: string; label: string }[]
> = {
  string: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "like", label: "contains (like)" },
    { value: "not-like", label: "does not contain" },
    { value: "begins-with", label: "begins with" },
    { value: "ends-with", label: "ends with" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  memo: [
    { value: "like", label: "contains" },
    { value: "not-like", label: "does not contain" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  integer: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "gt", label: "is greater than" },
    { value: "ge", label: "is greater than or equal" },
    { value: "lt", label: "is less than" },
    { value: "le", label: "is less than or equal" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  decimal: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "gt", label: "is greater than" },
    { value: "ge", label: "is greater than or equal" },
    { value: "lt", label: "is less than" },
    { value: "le", label: "is less than or equal" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  money: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "gt", label: "is greater than" },
    { value: "ge", label: "is greater than or equal" },
    { value: "lt", label: "is less than" },
    { value: "le", label: "is less than or equal" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  datetime: [
    { value: "on", label: "on" },
    { value: "on-or-before", label: "on or before" },
    { value: "on-or-after", label: "on or after" },
    { value: "last-x-days", label: "in last X days" },
    { value: "next-x-days", label: "in next X days" },
    { value: "older-than-x-days", label: "older than X days" },
    { value: "last-x-months", label: "in last X months" },
    { value: "today", label: "today" },
    { value: "yesterday", label: "yesterday" },
    { value: "tomorrow", label: "tomorrow" },
    { value: "this-week", label: "this week" },
    { value: "this-month", label: "this month" },
    { value: "this-year", label: "this year" },
    { value: "last-week", label: "last week" },
    { value: "last-month", label: "last month" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  boolean: [
    { value: "eq", label: "equals" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  picklist: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  lookup: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "eq-userid", label: "equals current user" },
    { value: "ne-userid", label: "not equal to current user" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
  unknown: [
    { value: "eq", label: "equals" },
    { value: "ne", label: "does not equal" },
    { value: "null", label: "is empty" },
    { value: "not-null", label: "is not empty" },
  ],
};

/**
 * Map raw AttributeType strings from Dataverse metadata to high-level buckets.
 */
export function bucketFor(rawType: string | undefined | null): AttributeBucket {
  if (!rawType) return "unknown";
  const t = rawType.toLowerCase();
  if (t === "string") return "string";
  if (t === "memo") return "memo";
  if (t === "integer" || t === "biginteger") return "integer";
  if (t === "decimal" || t === "double") return "decimal";
  if (t === "money") return "money";
  if (t === "datetime") return "datetime";
  if (t === "boolean") return "boolean";
  if (t === "picklist" || t === "state" || t === "status") return "picklist";
  if (t === "lookup" || t === "customer" || t === "owner") return "lookup";
  return "unknown";
}

/**
 * Escape XML special chars in attribute names and values.
 */
function xmlEscape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Turn the builder state into a well-formed FetchXML string.
 * Empty conditions (no attribute selected) are silently dropped.
 */
export function buildFetchXml(state: FetchXmlBuilderState): string {
  const validConditions = state.conditions.filter((c) => c.attribute && c.operator);
  const indent = "  ";
  const lines: string[] = [];

  const topAttr = state.top > 0 ? ` top="${state.top}"` : "";
  lines.push(`<fetch${topAttr}>`);
  lines.push(`${indent}<entity name="${xmlEscape(state.entity || "account")}">`);

  // attribute list — when empty, include just the primary id by convention.
  // (the flow will fetch full records via the SDK regardless of this)
  lines.push(`${indent}${indent}<all-attributes />`);

  if (validConditions.length > 0) {
    lines.push(`${indent}${indent}<filter type="${state.filterMode}">`);
    for (const c of validConditions) {
      const attr = xmlEscape(c.attribute);
      const op = xmlEscape(c.operator);
      if (NO_VALUE_OPERATORS.has(c.operator)) {
        lines.push(
          `${indent}${indent}${indent}<condition attribute="${attr}" operator="${op}" />`
        );
      } else {
        const val = xmlEscape(c.value ?? "");
        lines.push(
          `${indent}${indent}${indent}<condition attribute="${attr}" operator="${op}" value="${val}" />`
        );
      }
    }
    lines.push(`${indent}${indent}</filter>`);
  }

  if (state.sortAttribute) {
    lines.push(
      `${indent}${indent}<order attribute="${xmlEscape(state.sortAttribute)}" descending="${state.sortDescending ? "true" : "false"}" />`
    );
  }

  lines.push(`${indent}</entity>`);
  lines.push(`</fetch>`);

  return lines.join("\n");
}

/**
 * Create a fresh empty builder state for a target entity.
 */
export function emptyBuilderState(entity: string, top = 25): FetchXmlBuilderState {
  return {
    entity,
    filterMode: "and",
    conditions: [{ attribute: "", operator: "eq", value: "" }],
    sortAttribute: "",
    sortDescending: true,
    top,
  };
}

/**
 * Pretty-prints FetchXML with one element per line and two-space indentation.
 * Saved views store FetchXML on a single line, which is unreadable in the
 * query editor. Text content is preserved; already formatted XML is
 * re-flowed consistently.
 */
export function formatFetchXml(fetchXml: string): string {
  const compact = fetchXml.replace(/>\s+</g, "><").trim();
  if (!compact.startsWith("<")) return fetchXml;
  const tokens = compact.match(/<[^>]+>|[^<]+/g) ?? [];
  const lines: string[] = [];
  let depth = 0;
  for (const token of tokens) {
    if (!token.trim()) continue;
    const isClosing = /^<\//.test(token);
    const isSelfClosing = /\/>$/.test(token) || /^<\?/.test(token) || /^<!/.test(token);
    if (isClosing) depth = Math.max(0, depth - 1);
    lines.push(`${"  ".repeat(depth)}${token.trim()}`);
    if (token.startsWith("<") && !isClosing && !isSelfClosing) depth += 1;
  }
  return lines.join("\n");
}
