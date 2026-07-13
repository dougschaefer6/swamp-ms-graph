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

/**
 * `@dougschaefer/ms-graph-places` model — wraps the Microsoft Graph v1.0 Places
 * directory to enumerate and inspect room resources in the tenant.
 *
 * Part of the broad `@dougschaefer/ms-graph-*` family. `listRooms` returns every
 * room resource (displayName + emailAddress); `getRoom` returns detail for one
 * room — capacity, building, floor, A/V device names, accessibility, and tags.
 * The IARS meeting-agent uses the room mailbox addresses returned here as the
 * calendar-model input.
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permission: `Place.Read.All`
 * (admin consent required).
 *
 *   client_id:     ${{ vault.get(azure-graph, client_id) }}
 *   client_secret: ${{ vault.get(azure-graph, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-graph, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-places",
  version: "2026.07.13.4",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    roomList: {
      description:
        "Snapshot of all room resources in the tenant from /places/microsoft.graph.room. Each entry includes displayName and emailAddress.",
      schema: z.object({
        rooms: z.array(z.unknown()),
        count: z.number(),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "1d",
      garbageCollection: 3,
    },
    roomDetail: {
      description:
        "Detail for a single room resource from GET /places/{id-or-email}: displayName, emailAddress, capacity, building, floor, and A/V capabilities.",
      schema: z.object({
        id: z.string(),
        displayName: z.string(),
        emailAddress: z.string(),
        capacity: z.number(),
        building: z.string(),
        floorNumber: z.number(),
        isWheelChairAccessible: z.boolean(),
        phone: z.string(),
        nickname: z.string(),
        audioDeviceName: z.string(),
        videoDeviceName: z.string(),
        displayDeviceName: z.string(),
        tags: z.array(z.string()),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "7d",
      garbageCollection: 10,
    },
  },
  methods: {
    listRooms: {
      description:
        "GET /places/microsoft.graph.room — list all room resources in the tenant. Returns displayName + emailAddress for each room. Use the emailAddress as the roomEmail input for the calendar model. The method auto-follows @odata.nextLink. (Permission: Place.Read.All)",
      arguments: z.object({
        top: z.number().int().default(100).describe(
          "Max rooms to return per page (1-100). Paging is handled automatically.",
        ),
      }),
      execute: async (
        args: { top: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("Listing room resources from Graph /places");
        const url = `${GRAPH_BASE}/places/microsoft.graph.room?$top=${
          Math.min(100, Math.max(1, args.top))
        }&$select=displayName,emailAddress`;
        const rooms = await graphList(context.globalArgs, url);
        const handle = await context.writeResource("roomList", "list", {
          rooms,
          count: rooms.length,
          capturedAt: new Date().toISOString(),
        });
        context.logger.info("Found {n} room(s)", { n: rooms.length });
        return { dataHandles: [handle] };
      },
    },

    getRoom: {
      description:
        "GET /places/{idOrEmail} — fetch detail for a single room resource: capacity, building, floor, A/V device names, accessibility, and tags. Pass the room's emailAddress or object ID. (Permission: Place.Read.All)",
      arguments: z.object({
        idOrEmail: z.string().describe(
          "Room emailAddress (e.g. conf-b@example.com) or Graph object ID",
        ),
      }),
      execute: async (
        args: { idOrEmail: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("Fetching room detail for {id}", {
          id: args.idOrEmail,
        });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/places/${encodeURIComponent(args.idOrEmail)}`,
        );
        const d = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "roomDetail",
          slugify(args.idOrEmail),
          {
            id: str(d.id),
            displayName: str(d.displayName),
            emailAddress: str(d.emailAddress),
            capacity: Number(d.capacity ?? 0),
            building: str(d.building),
            floorNumber: Number(d.floorNumber ?? 0),
            isWheelChairAccessible: Boolean(d.isWheelChairAccessible),
            phone: str(d.phone),
            nickname: str(d.nickname),
            audioDeviceName: str(d.audioDeviceName),
            videoDeviceName: str(d.videoDeviceName),
            displayDeviceName: str(d.displayDeviceName),
            tags: Array.isArray(d.tags) ? d.tags.map(String) : [],
            capturedAt: new Date().toISOString(),
          },
        );
        context.logger.info("Room {name} fetched", {
          name: str(d.displayName),
        });
        return { dataHandles: [handle] };
      },
    },
  },
};
