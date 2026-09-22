import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { UserqueriesService } from "@/generated/services/UserqueriesService";

export interface UserView {
  id: string;
  name: string;
  description: string;
  returnedTypeCode: string;
  queryType: number;
  fetchXml: string;
}

function mapFromDataverse(raw: Record<string, unknown>): UserView {
  return {
    id: (raw.userqueryid as string) ?? "",
    name: (raw.name as string) ?? "(unnamed)",
    description: (raw.description as string) ?? "",
    returnedTypeCode: (raw.returnedtypecode as string) ?? "",
    queryType: (raw.querytype as number | undefined) ?? 0,
    fetchXml: (raw.fetchxml as string) ?? "",
  };
}

const QUERY_KEY = ["user-views"] as const;

/**
 * List the current user's personal views (userquery) for a given entity.
 * Dataverse returns only the views owned by the calling user — no need
 * to add an explicit owner filter.
 */
export function useUserViews(targetEntity: string | undefined) {
  return useQuery({
    queryKey: [...QUERY_KEY, targetEntity],
    enabled: !!targetEntity,
    queryFn: async (): Promise<UserView[]> => {
      if (!targetEntity) return [];
      const res = await UserqueriesService.getAll({
        filter: `returnedtypecode eq '${targetEntity}' and statecode eq 0`,
        top: 200,
      });
      const list = (res.data ?? []).map((r) =>
        mapFromDataverse(r as unknown as Record<string, unknown>)
      );
      list.sort((a, b) => a.name.localeCompare(b.name));
      return list;
    },
  });
}

export interface CreateUserViewInput {
  name: string;
  description?: string;
  returnedTypeCode: string;
  fetchXml: string;
}

/**
 * Save a FetchXML as a new personal view (userquery). The new view becomes
 * available everywhere the user can pick personal views — including model-
 * driven apps. Returns the new view's id.
 */
export function useCreateUserView() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (input: CreateUserViewInput): Promise<string> => {
      // querytype = 0 is the "user-defined view" type that shows up in
      // the model-driven app's personal views list.
      const payload = {
        name: input.name,
        description: input.description ?? "",
        returnedtypecode: input.returnedTypeCode,
        fetchxml: input.fetchXml,
        querytype: 0,
        // Minimal layoutxml — required by Dataverse so the view can be
        // listed. Shows just the primary name column.
        layoutxml: `<grid name="resultset" object="0" jump="name" select="1" icon="1" preview="1"><row name="result" id="${input.returnedTypeCode}id"><cell name="name" width="300" /></row></grid>`,
      } as never;
      const res = await UserqueriesService.create(payload);
      return (res.data as unknown as { userqueryid?: string })?.userqueryid ?? "";
    },
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: QUERY_KEY });
    },
  });
}
