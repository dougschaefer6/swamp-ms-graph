import { z } from "npm:zod@4.3.6";
import {
  type DataHandle,
  GRAPH_BASE,
  graphRequest,
  type MethodContext,
  MsGraphGlobalArgsSchema,
  slugify,
  str,
} from "./_client.ts";

const PresenceSchema = z
  .object({
    id: z.string(),
    availability: z.string().nullish(),
    activity: z.string().nullish(),
    capturedAt: z.iso.datetime(),
  })
  .passthrough();

/**
 * `@dougschaefer/ms-graph-presence` model — Microsoft Teams presence reads over
 * Microsoft Graph v1.0 with app-only client credentials.
 *
 * Part of the broad `@dougschaefer/ms-graph-*` family. `getPresence` returns a
 * user's current availability (Available, Busy, Away, DoNotDisturb, Offline) and
 * activity — useful for routing alerts to whoever is actually online, or for the
 * IARS room to know if an invited attendee is reachable.
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permission: `Presence.Read.All`;
 * the method returns HTTP 403 until an admin grants and consents that scope on
 * your app registration. See the README "Scopes to grant" section.
 *
 *   client_id:     ${{ vault.get(azure-graph, client_id) }}
 *   client_secret: ${{ vault.get(azure-graph, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-graph, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-presence",
  version: "2026.07.13.3",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    presence: {
      description:
        "A user's Teams presence from GET /users/{userId}/presence: availability and activity at capture time.",
      schema: PresenceSchema,
      lifetime: "2m",
      garbageCollection: 30,
    },
  },
  methods: {
    getPresence: {
      description:
        "GET /users/{userId}/presence — return a user's current Teams presence (availability + activity). (Permission: Presence.Read.All)",
      arguments: z.object({
        userId: z.string().describe(
          "User object id or userPrincipalName, e.g. user@example.com",
        ),
      }),
      execute: async (
        args: { userId: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("getPresence for {u}", { u: args.userId });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${encodeURIComponent(args.userId)}/presence`,
        );
        const p = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "presence",
          slugify(args.userId, "presence"),
          {
            id: str(p.id) || args.userId,
            availability: str(p.availability),
            activity: str(p.activity),
            capturedAt: new Date().toISOString(),
          },
        );
        context.logger.info("{u} is {a} / {act}", {
          u: args.userId,
          a: str(p.availability),
          act: str(p.activity),
        });
        return { dataHandles: [handle] };
      },
    },
  },
};
