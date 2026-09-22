import { useQuery } from "@tanstack/react-query";
import { SavedqueriesService } from "@/generated/services/SavedqueriesService";

export interface SystemView {
  id: string;
  name: string;
  description: string;
  returnedTypeCode: string;
  queryType: number;
  fetchXml: string;
  isDefault: boolean;
}

function mapFromDataverse(raw: Record<string, unknown>): SystemView {
  return {
    id: (raw.savedqueryid as string) ?? "",
    name: (raw.name as string) ?? "(unnamed)",
    description: (raw.description as string) ?? "",
    returnedTypeCode: (raw.returnedtypecode as string) ?? "",
    queryType: (raw.querytype as number | undefined) ?? 0,
    fetchXml: (raw.fetchxml as string) ?? "",
    isDefault: Boolean(raw.isdefault),
  };
}

/**
 * List system saved queries (savedquery) for a given target entity.
 *
 * Filters:
 *   - returnedtypecode = targetEntity
 *   - querytype = 0 (main view) or 1 (advanced find) — these are the kinds
 *     of views a user would intentionally name and save. Skips
 *     querytype 2 (lookup views), 4 (mobile), etc.
 *   - statecode eq 0 (active)
 */
export function useSystemViews(targetEntity: string | undefined) {
  return useQuery({
    queryKey: ["system-views", targetEntity],
    enabled: !!targetEntity,
    queryFn: async (): Promise<SystemView[]> => {
      if (!targetEntity) return [];
      const res = await SavedqueriesService.getAll({
        filter: `returnedtypecode eq '${targetEntity}' and (querytype eq 0 or querytype eq 1) and statecode eq 0`,
        top: 200,
      });
      const list = (res.data ?? []).map((r) =>
        mapFromDataverse(r as unknown as Record<string, unknown>)
      );
      // Sort: defaults first, then alphabetically by name
      list.sort((a, b) => {
        if (a.isDefault !== b.isDefault) return a.isDefault ? -1 : 1;
        return a.name.localeCompare(b.name);
      });
      return list;
    },
  });
}
