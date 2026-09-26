import { useEffect, useMemo, useState } from "react";
import {
  ArrowDown,
  ArrowUp,
  Calendar,
  Database,
  Filter as FilterIcon,
  Hash,
  Loader2,
  Plus,
  Sparkles,
  ToggleLeft,
  Type,
  Wand2,
  X,
} from "lucide-react";
import { useEntityAttributes, type EntityAttribute } from "@/hooks/useEntityAttributes";
import {
  bucketFor,
  buildFetchXml,
  emptyBuilderState,
  NO_VALUE_OPERATORS,
  OPERATORS_BY_BUCKET,
  type AttributeBucket,
  type FilterCondition,
  type FetchXmlBuilderState,
} from "@/lib/fetchxml";
import { cn } from "@/lib/utils";

/* ─── helpers ────────────────────────────────────────────────────── */

function renderBucketIcon(bucket: AttributeBucket) {
  const Icon = bucketIcon(bucket);
  return <Icon className="h-3 w-3 shrink-0 text-muted-foreground" strokeWidth={1.75} />;
}

function bucketIcon(bucket: AttributeBucket) {
  switch (bucket) {
    case "string":
    case "memo":
      return Type;
    case "integer":
    case "decimal":
    case "money":
      return Hash;
    case "datetime":
      return Calendar;
    case "boolean":
      return ToggleLeft;
    case "picklist":
    case "lookup":
      return Database;
    default:
      return Type;
  }
}

function valueInputType(bucket: AttributeBucket, operator: string): "text" | "number" | "date" {
  if (operator.includes("x-days") || operator.includes("x-months") || operator.includes("x-hours") || operator.includes("x-weeks") || operator.includes("x-years")) {
    return "number";
  }
  if (bucket === "datetime") return "date";
  if (bucket === "integer" || bucket === "decimal" || bucket === "money") return "number";
  return "text";
}

/* ─── Attribute combobox (datalist-backed) ──────────────────────── */

function AttributeCombobox({
  value,
  attributes,
  onChange,
  loading,
  id,
}: {
  value: string;
  attributes: EntityAttribute[];
  onChange: (logicalName: string) => void;
  loading: boolean;
  id: string;
}) {
  // We use a datalist so users get autocomplete but can still type any
  // attribute (including custom columns the metadata might not return).
  const visible = attributes.filter((a) => !a.isSystem);
  const hidden = attributes.filter((a) => a.isSystem);
  return (
    <>
      <input
        list={id}
        value={value}
        onChange={(e) => onChange(e.target.value.trim())}
        placeholder={loading ? "Loading…" : "Pick an attribute…"}
        spellCheck={false}
        className="h-8 w-full rounded-md border border-border bg-background px-2.5 font-mono text-[11px] text-foreground focus:border-primary/50 focus:outline-none focus:ring-1 focus:ring-primary/30"
      />
      <datalist id={id}>
        {visible.map((a) => (
          <option key={a.logicalName} value={a.logicalName}>
            {a.displayName} · {a.rawType}
          </option>
        ))}
        {hidden.length > 0 && (
          <option disabled value="">
            — system fields below —
          </option>
        )}
        {hidden.map((a) => (
          <option key={a.logicalName} value={a.logicalName}>
            {a.displayName} · {a.rawType}
          </option>
        ))}
      </datalist>
    </>
  );
}

/* ─── Single filter condition row ───────────────────────────────── */

function ConditionRow({
  condition,
  attributes,
  attributesLoading,
  onChange,
  onRemove,
  idx,
}: {
  condition: FilterCondition;
  attributes: EntityAttribute[];
  attributesLoading: boolean;
  onChange: (next: FilterCondition) => void;
  onRemove: () => void;
  idx: number;
}) {
  // Find the bucket for the picked attribute (defaults to unknown).
  const meta = attributes.find((a) => a.logicalName === condition.attribute);
  const bucket: AttributeBucket = meta ? meta.bucket : bucketFor(undefined);
  const operators = OPERATORS_BY_BUCKET[bucket];
  const needsValue = !NO_VALUE_OPERATORS.has(condition.operator);
  const inputType = valueInputType(bucket, condition.operator);

  // If the picked operator isn't valid for this bucket, reset to the first.
  useEffect(() => {
    if (!operators.find((op) => op.value === condition.operator)) {
      onChange({ ...condition, operator: operators[0]?.value ?? "eq" });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [bucket]);

  return (
    <div className="grid grid-cols-[20px_minmax(140px,_1.2fr)_minmax(140px,_1fr)_minmax(120px,_1fr)_28px] items-center gap-2 rounded-md border border-border bg-card px-2 py-1.5">
      <span className="grid h-6 w-5 place-items-center text-[10px] tabular-nums text-muted-foreground">
        {(idx + 1).toString().padStart(2, "0")}
      </span>
      <div className="flex items-center gap-1.5">
        {renderBucketIcon(bucket)}
        <AttributeCombobox
          id={`attrs-${idx}`}
          value={condition.attribute}
          attributes={attributes}
          loading={attributesLoading}
          onChange={(v) => onChange({ ...condition, attribute: v })}
        />
      </div>
      <select
        value={condition.operator}
        onChange={(e) => onChange({ ...condition, operator: e.target.value })}
        className="h-8 w-full rounded-md border border-border bg-background px-2 text-[11px] text-foreground focus:border-primary/50 focus:outline-none focus:ring-1 focus:ring-primary/30"
      >
        {operators.map((op) => (
          <option key={op.value} value={op.value}>
            {op.label}
          </option>
        ))}
      </select>
      {needsValue ? (
        <input
          type={inputType}
          value={condition.value}
          onChange={(e) => onChange({ ...condition, value: e.target.value })}
          placeholder="value"
          spellCheck={false}
          className="h-8 w-full rounded-md border border-border bg-background px-2 font-mono text-[11px] text-foreground focus:border-primary/50 focus:outline-none focus:ring-1 focus:ring-primary/30"
        />
      ) : (
        <span className="grid h-8 place-items-center rounded-md bg-muted/50 px-2 text-[10px] italic text-muted-foreground">
          (no value needed)
        </span>
      )}
      <button
        type="button"
        onClick={onRemove}
        aria-label="Remove condition"
        className="grid h-6 w-6 place-items-center rounded-md text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
      >
        <X className="h-3.5 w-3.5" strokeWidth={2} />
      </button>
    </div>
  );
}

/* ─── Main builder component ────────────────────────────────────── */

export interface VisualFilterBuilderProps {
  targetEntity: string;
  /** Initial state — if not provided, an empty state is created for the entity. */
  initialState?: FetchXmlBuilderState;
  /** Called with the rendered FetchXML when the user clicks Generate. */
  onGenerate: (fetchXml: string, state: FetchXmlBuilderState) => void;
}

export function VisualFilterBuilder({
  targetEntity,
  initialState,
  onGenerate,
}: VisualFilterBuilderProps) {
  const [state, setState] = useState<FetchXmlBuilderState>(
    () => initialState ?? emptyBuilderState(targetEntity)
  );

  // Re-initialise state when the target entity changes from outside.
  const [trackedEntity, setTrackedEntity] = useState(targetEntity);
  if (trackedEntity !== targetEntity) {
    setTrackedEntity(targetEntity);
    setState(emptyBuilderState(targetEntity));
  }

  const attributesQuery = useEntityAttributes(targetEntity);
  const attributes = attributesQuery.data ?? [];
  const livePreviewXml = useMemo(() => buildFetchXml(state), [state]);

  function patch<K extends keyof FetchXmlBuilderState>(
    key: K,
    value: FetchXmlBuilderState[K]
  ) {
    setState((s) => ({ ...s, [key]: value }));
  }

  function updateCondition(i: number, next: FilterCondition) {
    setState((s) => ({
      ...s,
      conditions: s.conditions.map((c, idx) => (idx === i ? next : c)),
    }));
  }

  function removeCondition(i: number) {
    setState((s) => ({
      ...s,
      conditions: s.conditions.filter((_, idx) => idx !== i),
    }));
  }

  function addCondition() {
    setState((s) => ({
      ...s,
      conditions: [...s.conditions, { attribute: "", operator: "eq", value: "" }],
    }));
  }

  function handleGenerate() {
    onGenerate(livePreviewXml, state);
  }

  const validConditionsCount = state.conditions.filter(
    (c) => c.attribute && c.operator
  ).length;

  return (
    <div className="space-y-4">
      {!targetEntity && (
        <p className="rounded-md border border-dashed border-border px-3 py-2 text-[11px] text-muted-foreground">
          Pick a target entity above first.
        </p>
      )}

      {targetEntity && attributes.length === 0 && !attributesQuery.isFetching && (
        <div className="rounded-md border border-status-watch/30 bg-status-watch/5 px-3 py-2 text-[11px] text-status-watch">
          Couldn't load attribute metadata for{" "}
          <code className="rounded bg-muted px-1 font-mono text-[10px]">
            {targetEntity}
          </code>
          . You can still type attribute logical names manually — the FetchXML will
          be valid as long as the names match real columns.
        </div>
      )}

      {/* Filter section */}
      <section>
        <div className="mb-2 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <FilterIcon className="h-3.5 w-3.5 text-primary" strokeWidth={2} />
            <span className="text-[11px] font-semibold uppercase tracking-wider text-foreground">
              Filter
            </span>
            <span className="text-[10px] text-muted-foreground">
              {validConditionsCount} condition
              {validConditionsCount === 1 ? "" : "s"}
            </span>
            {attributesQuery.isFetching && (
              <Loader2 className="h-3 w-3 animate-spin text-muted-foreground" />
            )}
          </div>
          <div className="inline-flex rounded-md border border-border bg-card p-0.5">
            <button
              type="button"
              onClick={() => patch("filterMode", "and")}
              className={cn(
                "rounded px-2 py-0.5 text-[10px] font-medium transition-colors",
                state.filterMode === "and"
                  ? "bg-primary text-primary-foreground"
                  : "text-muted-foreground hover:text-foreground"
              )}
            >
              ALL (and)
            </button>
            <button
              type="button"
              onClick={() => patch("filterMode", "or")}
              className={cn(
                "rounded px-2 py-0.5 text-[10px] font-medium transition-colors",
                state.filterMode === "or"
                  ? "bg-primary text-primary-foreground"
                  : "text-muted-foreground hover:text-foreground"
              )}
            >
              ANY (or)
            </button>
          </div>
        </div>

        <div className="space-y-1.5">
          {state.conditions.map((c, i) => (
            <ConditionRow
              key={i}
              idx={i}
              condition={c}
              attributes={attributes}
              attributesLoading={attributesQuery.isFetching}
              onChange={(next) => updateCondition(i, next)}
              onRemove={() => removeCondition(i)}
            />
          ))}
        </div>

        <button
          type="button"
          onClick={addCondition}
          className="mt-2 inline-flex items-center gap-1.5 rounded-md border border-dashed border-border px-2.5 py-1 text-[11px] font-medium text-muted-foreground transition-colors hover:border-primary/40 hover:bg-primary/5 hover:text-primary"
        >
          <Plus className="h-3 w-3" strokeWidth={2.25} />
          Add condition
        </button>
      </section>

      {/* Sort + Top */}
      <section className="grid grid-cols-[2fr_auto_1fr] items-end gap-3">
        <div>
          <label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
            Sort by
          </label>
          <div className="mt-1.5 flex gap-1.5">
            <input
              list="sort-attr-list"
              value={state.sortAttribute}
              onChange={(e) => patch("sortAttribute", e.target.value.trim())}
              placeholder="(no sort)"
              spellCheck={false}
              className="h-8 flex-1 rounded-md border border-border bg-background px-2.5 font-mono text-[11px] text-foreground focus:border-primary/50 focus:outline-none focus:ring-1 focus:ring-primary/30"
            />
            <datalist id="sort-attr-list">
              {attributes
                .filter((a) => !a.isSystem)
                .map((a) => (
                  <option key={a.logicalName} value={a.logicalName}>
                    {a.displayName}
                  </option>
                ))}
            </datalist>
          </div>
        </div>
        <button
          type="button"
          onClick={() => patch("sortDescending", !state.sortDescending)}
          disabled={!state.sortAttribute}
          title={state.sortDescending ? "Descending" : "Ascending"}
          className="grid h-8 w-8 place-items-center rounded-md border border-border bg-card text-foreground transition-colors hover:bg-muted disabled:opacity-50"
        >
          {state.sortDescending ? (
            <ArrowDown className="h-3.5 w-3.5" strokeWidth={2} />
          ) : (
            <ArrowUp className="h-3.5 w-3.5" strokeWidth={2} />
          )}
        </button>
        <div>
          <label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
            Top N
          </label>
          <input
            type="number"
            min={1}
            max={5000}
            value={state.top}
            onChange={(e) => patch("top", parseInt(e.target.value, 10) || 25)}
            className="mt-1.5 h-8 w-full rounded-md border border-border bg-background px-2.5 text-[11px] tabular-nums text-foreground focus:border-primary/50 focus:outline-none focus:ring-1 focus:ring-primary/30"
          />
        </div>
      </section>

      {/* Live FetchXML preview */}
      <section>
        <div className="mb-1.5 flex items-center justify-between">
          <span className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
            Generated FetchXML
          </span>
          <span className="text-[10px] text-muted-foreground">
            {livePreviewXml.length} chars · live
          </span>
        </div>
        <pre className="max-h-48 overflow-auto rounded-md border border-border bg-muted/40 px-3 py-2 font-mono text-[11px] leading-relaxed text-foreground/85">
          {livePreviewXml}
        </pre>
      </section>

      {/* Action */}
      <div className="flex items-center justify-end gap-2">
        <button
          type="button"
          onClick={handleGenerate}
          disabled={!targetEntity}
          className="inline-flex items-center gap-1.5 rounded-md bg-primary px-3 py-1.5 text-[12px] font-medium text-primary-foreground shadow-sm transition-all hover:bg-primary/90 active:scale-95 disabled:opacity-60"
        >
          <Wand2 className="h-3.5 w-3.5" strokeWidth={2.25} />
          Use this FetchXML
        </button>
      </div>

      <p className="text-[10px] text-muted-foreground">
        <Sparkles className="mr-1 inline h-2.5 w-2.5" strokeWidth={2} />
        For joins, link-entity, aggregates or nested filter groups, switch to the{" "}
        <strong>Paste FetchXML</strong> tab and edit by hand.
      </p>
    </div>
  );
}

