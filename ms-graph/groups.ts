import { z } from "npm:zod@4.3.6";
import {
  type DataHandle,
  GRAPH_BASE,
  graphList,
  graphRequest,
  type MethodContext,
  MsGraphGlobalArgsSchema,
  slugify,
  str,
} from "./_client.ts";

/** Standard set of group properties selected on every Graph group read. */
const GROUP_SELECT =
  "id,displayName,description,mail,mailEnabled,securityEnabled," +
  "groupTypes,visibility,membershipRule";

const GroupSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
    mail: z.string().nullish(),
    mailEnabled: z.boolean().optional(),
    securityEnabled: z.boolean().optional(),
    groupTypes: z.array(z.string()).optional(),
    visibility: z.string().nullish(),
    membershipRule: z.string().nullish(),
  })
  .passthrough();

const MemberSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    userPrincipalName: z.string().nullish(),
    mail: z.string().nullish(),
    "@odata.type": z.string().nullish(),
  })
  .passthrough();

/**
 * `@dougschaefer/ms-graph-groups` model — Entra ID group directory reads over
 * Microsoft Graph v1.0 with app-only client credentials.
 *
 * Part of the broad `@dougschaefer/ms-graph-*` family. `list` enumerates groups
 * (optionally OData-filtered), `get` resolves one group by object id, and
 * `listMembers` returns a group's members (users, devices, nested groups).
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permission: `Group.Read.All`.
 * As of this writing the Swamp Entra app has NOT been consented `Group.Read.All`,
 * so these methods are built but will return HTTP 403 until an admin grants and
 * consents the scope. See the README "Scopes to grant" section.
 *
 *   client_id:     ${{ vault.get(azure-asei, client_id) }}
 *   client_secret: ${{ vault.get(azure-asei, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-asei, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-groups",
  version: "2026.06.29.1",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    group: {
      description:
        "An Entra ID group from GET /groups/{id}, selected to the standard fields (displayName, mail, securityEnabled, groupTypes, membershipRule, etc.).",
      schema: GroupSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    member: {
      description:
        "A member of a group (user, device, or nested group) from /groups/{id}/members.",
      schema: MemberSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
  },
  methods: {
    list: {
      description:
        "GET /groups — enumerate directory groups, selected to the standard fields. Optionally narrow with an OData $filter and cap the total with maxItems. Auto-follows @odata.nextLink. (Permission: Group.Read.All)",
      arguments: z.object({
        filter: z.string().optional().describe(
          'OData $filter, e.g. "securityEnabled eq true" or "startswith(displayName,\'AV\')"',
        ),
        maxItems: z.number().int().default(200).describe(
          "Hard cap on total groups collected across pages (default 200).",
        ),
      }),
      execute: async (
        args: { filter?: string; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        let url = `${GRAPH_BASE}/groups?$select=${GROUP_SELECT}&$top=100`;
        if (args.filter) url += `&$filter=${encodeURIComponent(args.filter)}`;
        const groups = await graphList(context.globalArgs, url, {
          maxItems: args.maxItems,
        });
        context.logger.info("Found {count} group(s)", { count: groups.length });
        const handles: DataHandle[] = [];
        for (const g of groups) {
          const go = (g ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "group",
              slugify(str(go.id), "group"),
              go,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    get: {
      description:
        "GET /groups/{id} — resolve one group by Entra object id, selected to the standard fields. (Permission: Group.Read.All)",
      arguments: z.object({
        id: z.string().describe("Entra group object id (GUID)"),
      }),
      execute: async (
        args: { id: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("Resolving group {id}", { id: args.id });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/groups/${encodeURIComponent(args.id)}` +
            `?$select=${GROUP_SELECT}`,
        );
        const g = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "group",
          slugify(str(g.id), "group"),
          g,
        );
        context.logger.info("Resolved group {name}", {
          name: str(g.displayName),
        });
        return { dataHandles: [handle] };
      },
    },

    listMembers: {
      description:
        "GET /groups/{id}/members — list the members of a group (users, devices, nested groups). Auto-follows @odata.nextLink. (Permission: Group.Read.All)",
      arguments: z.object({
        id: z.string().describe("Entra group object id (GUID)"),
        maxItems: z.number().int().default(500).describe(
          "Hard cap on total members collected across pages (default 500).",
        ),
      }),
      execute: async (
        args: { id: string; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const members = await graphList(
          context.globalArgs,
          `${GRAPH_BASE}/groups/${
            encodeURIComponent(args.id)
          }/members?$top=100`,
          { maxItems: args.maxItems },
        );
        context.logger.info("Group {id} has {count} member(s)", {
          id: args.id,
          count: members.length,
        });
        const handles: DataHandle[] = [];
        for (const m of members) {
          const mo = (m ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "member",
              slugify(str(mo.id), "member"),
              mo,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },
  },
};
