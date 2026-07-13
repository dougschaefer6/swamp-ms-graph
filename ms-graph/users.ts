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

/** Standard set of user properties selected on every Graph user read. */
const USER_SELECT =
  "id,displayName,userPrincipalName,givenName,surname,jobTitle,department," +
  "mail,mobilePhone,officeLocation,businessPhones,preferredLanguage,accountEnabled";

const UserSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    userPrincipalName: z.string().nullish(),
    givenName: z.string().nullish(),
    surname: z.string().nullish(),
    jobTitle: z.string().nullish(),
    department: z.string().nullish(),
    mail: z.string().nullish(),
    mobilePhone: z.string().nullish(),
    officeLocation: z.string().nullish(),
    businessPhones: z.array(z.string()).optional(),
    preferredLanguage: z.string().nullish(),
    accountEnabled: z.boolean().optional(),
  })
  .passthrough();

const DirectoryObjectSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    "@odata.type": z.string().nullish(),
  })
  .passthrough();

/**
 * `@dougschaefer/ms-graph-users` model — Entra ID (Azure AD) user directory reads
 * over Microsoft Graph v1.0 with app-only client credentials.
 *
 * This model absorbs the lookup half of the older `@dougschaefer/azure-ad-user`
 * type (which wrapped the `az ad user` CLI under an interactive login). Here the
 * read path is pure Graph REST against the app registration, so it runs unattended
 * inside workflows with no `az login` session. `get` resolves one user by object id
 * or UPN — the IARS reverse-lookup signal — and writes the full user object under a
 * slug of the object id (the user fields land under the data instance's `content`).
 * `list` enumerates the directory (optionally narrowed by an OData `$filter`),
 * `getManager` returns a user's manager, and `memberOf` returns the groups and
 * directory roles a user belongs to for access review.
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permission: `User.Read.All`
 * (admin-consented and live-verified in the tenant). `getManager` and
 * `memberOf` also resolve under `User.Read.All`.
 *
 *   client_id:     ${{ vault.get(azure-graph, client_id) }}
 *   client_secret: ${{ vault.get(azure-graph, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-graph, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-users",
  version: "2026.07.13.3",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    user: {
      description:
        "An Entra ID user from GET /users/{id}, selected to the standard profile fields. Written under a slug of the object id; the user fields land under the data instance's content.",
      schema: UserSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    membership: {
      description:
        "A directory object (group or directory role) a user is a member of, from /users/{id}/memberOf.",
      schema: DirectoryObjectSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
  },
  methods: {
    get: {
      description:
        "GET /users/{idOrUpn} — resolve one user by Entra object id or userPrincipalName, selected to the standard profile fields (displayName, department, jobTitle, officeLocation, mail, etc.). Writes the user under a slug of the object id. This is the IARS reverse-lookup signal — pass the object id. (Permission: User.Read.All)",
      arguments: z.object({
        idOrUpn: z.string().describe(
          "Entra object id (GUID) or userPrincipalName (user@example.com)",
        ),
      }),
      execute: async (
        args: { idOrUpn: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("Resolving user {id}", { id: args.idOrUpn });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${encodeURIComponent(args.idOrUpn)}` +
            `?$select=${USER_SELECT}`,
        );
        const u = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "user",
          slugify(str(u.id) || args.idOrUpn, "user"),
          u,
        );
        context.logger.info("Resolved {name} ({upn})", {
          name: str(u.displayName),
          upn: str(u.userPrincipalName),
        });
        return { dataHandles: [handle] };
      },
    },

    list: {
      description:
        "GET /users — enumerate directory users, selected to the standard profile fields. Optionally narrow with an OData $filter and cap the total with maxItems. Auto-follows @odata.nextLink. (Permission: User.Read.All)",
      arguments: z.object({
        filter: z.string().optional().describe(
          'OData $filter, e.g. "startswith(displayName,\'A\')" or "accountEnabled eq true"',
        ),
        maxItems: z.number().int().default(200).describe(
          "Hard cap on total users collected across pages (default 200).",
        ),
      }),
      execute: async (
        args: { filter?: string; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        let url = `${GRAPH_BASE}/users?$select=${USER_SELECT}&$top=100`;
        if (args.filter) {
          url += `&$filter=${encodeURIComponent(args.filter)}`;
        }
        const users = await graphList(context.globalArgs, url, {
          maxItems: args.maxItems,
        });
        context.logger.info("Found {count} user(s)", { count: users.length });
        const handles: DataHandle[] = [];
        for (const u of users) {
          const uo = (u ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "user",
              slugify(str(uo.id), "user"),
              uo,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    getManager: {
      description:
        "GET /users/{idOrUpn}/manager — return a user's manager as a user record. Throws if the user has no manager assigned. (Permission: User.Read.All)",
      arguments: z.object({
        idOrUpn: z.string().describe(
          "Entra object id (GUID) or userPrincipalName of the user whose manager to fetch",
        ),
      }),
      execute: async (
        args: { idOrUpn: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("Fetching manager of {id}", { id: args.idOrUpn });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${encodeURIComponent(args.idOrUpn)}/manager` +
            `?$select=${USER_SELECT}`,
        );
        const m = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "user",
          slugify(str(m.id), "manager"),
          m,
        );
        context.logger.info("Manager is {name}", { name: str(m.displayName) });
        return { dataHandles: [handle] };
      },
    },

    memberOf: {
      description:
        "GET /users/{idOrUpn}/memberOf — list the groups and directory roles a user is a member of, for access review. Auto-follows @odata.nextLink. (Permission: User.Read.All)",
      arguments: z.object({
        idOrUpn: z.string().describe(
          "Entra object id (GUID) or userPrincipalName of the user",
        ),
      }),
      execute: async (
        args: { idOrUpn: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const objs = await graphList(
          context.globalArgs,
          `${GRAPH_BASE}/users/${
            encodeURIComponent(args.idOrUpn)
          }/memberOf?$top=100`,
        );
        context.logger.info("User {id} is a member of {count} object(s)", {
          id: args.idOrUpn,
          count: objs.length,
        });
        const handles: DataHandle[] = [];
        for (const o of objs) {
          const oo = (o ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "membership",
              slugify(str(oo.id), "member"),
              oo,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },
  },
};
