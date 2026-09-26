import { useQuery } from "@tanstack/react-query";
import { isPowerAppsRuntime, runFetchXml } from "@/hooks/useDataverseCatalog";
import { getTargetEntityIdentity } from "@/lib/summaryConfiguration";
import type { StudioAccount, StudioConfiguration } from "@/hooks/useStudioRuntime";

export interface OutputSummary {
  id: string;
  name: string;
  summary: string;
  generatedOn: string | null;
}

function xmlAttribute(value: string) {
  return value.replace(/[^a-z0-9_]/gi, "");
}

export function buildOutputSummariesFetchXml(
  configuration: Pick<
    StudioConfiguration,
    "outputEntity" | "outputField" | "entityIdField" | "entity"
  >,
  primaryNameAttribute: string,
) {
  const entity = xmlAttribute(configuration.outputEntity);
  const outputField = xmlAttribute(configuration.outputField);
  const idField =
    configuration.outputEntity === configuration.entity && configuration.entityIdField
      ? xmlAttribute(configuration.entityIdField)
      : getTargetEntityIdentity(configuration.outputEntity).idField;
  const nameField = xmlAttribute(primaryNameAttribute || "name");
  const generatedOn =
    entity === "account" ? '<attribute name="csp_aisummarygeneratedon" />' : "";
  return (
    `<fetch><entity name="${entity}">` +
    `<attribute name="${idField}" /><attribute name="${nameField}" />` +
    `<attribute name="${outputField}" /><attribute name="modifiedon" />${generatedOn}` +
    `<filter><condition attribute="${outputField}" operator="not-null" /></filter>` +
    `<order attribute="modifiedon" descending="true" /></entity></fetch>`
  );
}

/**
 * Records that already carry a generated summary for a configuration:
 * rows of the destination table whose destination column is not empty.
 */
export function useOutputSummaries(
  configuration: StudioConfiguration | undefined,
  primaryNameAttribute: string,
  localAccounts: StudioAccount[],
  limit = 50,
) {
  return useQuery({
    queryKey: [
      "output-summaries",
      configuration?.id,
      configuration?.outputEntity,
      configuration?.outputField,
      primaryNameAttribute,
      limit,
    ],
    enabled: Boolean(configuration),
    staleTime: 15_000,
    queryFn: async (): Promise<OutputSummary[]> => {
      if (!configuration) return [];
      if (!isPowerAppsRuntime()) {
        return localAccounts.map((account) => ({
          id: account.id,
          name: account.name,
          summary: account.summary,
          generatedOn: account.generatedOn,
        }));
      }
      const entitySetName =
        configuration.outputEntitySetName ||
        getTargetEntityIdentity(configuration.outputEntity).entitySet;
      const idField =
        configuration.outputEntity === configuration.entity && configuration.entityIdField
          ? configuration.entityIdField
          : getTargetEntityIdentity(configuration.outputEntity).idField;
      const nameField = primaryNameAttribute || "name";
      const records = await runFetchXml(
        entitySetName,
        buildOutputSummariesFetchXml(configuration, nameField),
        limit,
        "Dataverse could not load the generated summaries.",
      );
      return records.map((raw) => ({
        id: String(raw[idField] ?? ""),
        name: String(raw[nameField] ?? raw.name ?? "Untitled record"),
        summary: String(raw[configuration.outputField] ?? ""),
        generatedOn: raw.csp_aisummarygeneratedon
          ? String(raw.csp_aisummarygeneratedon)
          : raw.modifiedon
            ? String(raw.modifiedon)
            : null,
      }));
    },
  });
}
