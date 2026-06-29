import { z } from "npm:zod@4.3.6";
import {
  type DataHandle,
  GRAPH_BASE,
  graphList,
  type MethodContext,
  MsGraphGlobalArgsSchema,
  slugify,
  str,
} from "./_client.ts";

const ChatSchema = z
  .object({
    id: z.string(),
    topic: z.string().nullish(),
    chatType: z.string().nullish(),
    createdDateTime: z.string().nullish(),
    lastUpdatedDateTime: z.string().nullish(),
  })
  .passthrough();

const TeamSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
  })
  .passthrough();

const ChannelSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
    membershipType: z.string().nullish(),
  })
  .passthrough();

/**
 * `@dougschaefer/ms-graph-teams` model — Microsoft Teams chat, team, channel, and
 * message reads over Microsoft Graph v1.0 with app-only client credentials.
 *
 * Part of the broad `@dougschaefer/ms-graph-*` family. `listChats` and
 * `listJoinedTeams` enumerate a user's chats and teams; `listChannels` lists a
 * team's channels; `listMessages` reads messages from either a chat or a channel.
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permissions: `Chat.Read.All`
 * (chats and chat messages) and `ChannelMessage.Read.All` (channel messages); team
 * and channel enumeration also needs `Team.ReadBasic.All` / `Channel.ReadBasic.All`.
 * As of this writing the Swamp Entra app has NOT been consented any of these,
 * so these methods are built but will return HTTP 403 until an admin grants and
 * consents the scopes. See the README "Scopes to grant" section.
 *
 *   client_id:     ${{ vault.get(azure-asei, client_id) }}
 *   client_secret: ${{ vault.get(azure-asei, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-asei, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-teams",
  version: "2026.06.29.1",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    chat: {
      description: "A Teams chat from /users/{userId}/chats.",
      schema: ChatSchema,
      lifetime: "1h",
      garbageCollection: 20,
    },
    team: {
      description:
        "A team the user has joined, from /users/{userId}/joinedTeams.",
      schema: TeamSchema,
      lifetime: "1d",
      garbageCollection: 10,
    },
    channel: {
      description: "A channel of a team, from /teams/{teamId}/channels.",
      schema: ChannelSchema,
      lifetime: "1d",
      garbageCollection: 10,
    },
    messageList: {
      description:
        "Snapshot of messages from a chat or channel, from /chats/{id}/messages or /teams/{teamId}/channels/{channelId}/messages.",
      schema: z.object({
        source: z.string(),
        messages: z.array(z.unknown()),
        count: z.number(),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "5m",
      garbageCollection: 20,
    },
  },
  methods: {
    listChats: {
      description:
        "GET /users/{userId}/chats — list a user's Teams chats. Auto-follows @odata.nextLink. (Permission: Chat.Read.All)",
      arguments: z.object({
        userId: z.string().describe(
          "User userPrincipalName or object id, e.g. user@example.com",
        ),
        maxItems: z.number().int().default(100).describe(
          "Hard cap on total chats collected across pages (default 100).",
        ),
      }),
      execute: async (
        args: { userId: string; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const chats = await graphList(
          context.globalArgs,
          `${GRAPH_BASE}/users/${
            encodeURIComponent(args.userId)
          }/chats?$top=50`,
          { maxItems: args.maxItems },
        );
        context.logger.info("Found {n} chat(s) for {u}", {
          n: chats.length,
          u: args.userId,
        });
        const handles: DataHandle[] = [];
        for (const c of chats) {
          const co = (c ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "chat",
              slugify(str(co.id), "chat"),
              co,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    listJoinedTeams: {
      description:
        "GET /users/{userId}/joinedTeams — list the teams a user has joined. Auto-follows @odata.nextLink. (Permission: Team.ReadBasic.All)",
      arguments: z.object({
        userId: z.string().describe(
          "User userPrincipalName or object id, e.g. user@example.com",
        ),
      }),
      execute: async (
        args: { userId: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const teams = await graphList(
          context.globalArgs,
          `${GRAPH_BASE}/users/${
            encodeURIComponent(args.userId)
          }/joinedTeams?$top=100`,
        );
        context.logger.info("Found {n} team(s) for {u}", {
          n: teams.length,
          u: args.userId,
        });
        const handles: DataHandle[] = [];
        for (const t of teams) {
          const to = (t ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "team",
              slugify(str(to.id), "team"),
              to,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    listChannels: {
      description:
        "GET /teams/{teamId}/channels — list a team's channels. Auto-follows @odata.nextLink. (Permission: Channel.ReadBasic.All)",
      arguments: z.object({
        teamId: z.string().describe("Team (group) object id"),
      }),
      execute: async (
        args: { teamId: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const channels = await graphList(
          context.globalArgs,
          `${GRAPH_BASE}/teams/${encodeURIComponent(args.teamId)}/channels`,
        );
        context.logger.info("Found {n} channel(s) for team {t}", {
          n: channels.length,
          t: args.teamId,
        });
        const handles: DataHandle[] = [];
        for (const c of channels) {
          const co = (c ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "channel",
              slugify(str(co.id), "channel"),
              co,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    listMessages: {
      description:
        "List messages from a chat (pass chatId) or a channel (pass teamId + channelId). Auto-follows @odata.nextLink. " +
        "(Permissions: Chat.Read.All for chat messages, ChannelMessage.Read.All for channel messages)",
      arguments: z.object({
        chatId: z.string().optional().describe(
          "Chat id — provide for chat messages (mutually exclusive with teamId/channelId)",
        ),
        teamId: z.string().optional().describe(
          "Team id — provide with channelId for channel messages",
        ),
        channelId: z.string().optional().describe(
          "Channel id — provide with teamId for channel messages",
        ),
        maxItems: z.number().int().default(50).describe(
          "Hard cap on total messages collected across pages (default 50).",
        ),
      }),
      execute: async (
        args: {
          chatId?: string;
          teamId?: string;
          channelId?: string;
          maxItems: number;
        },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        let url: string;
        let source: string;
        if (args.chatId) {
          url = `${GRAPH_BASE}/chats/${
            encodeURIComponent(args.chatId)
          }/messages?$top=50`;
          source = `chat-${args.chatId}`;
        } else if (args.teamId && args.channelId) {
          url = `${GRAPH_BASE}/teams/${
            encodeURIComponent(args.teamId)
          }/channels/${encodeURIComponent(args.channelId)}/messages?$top=50`;
          source = `channel-${args.channelId}`;
        } else {
          throw new Error(
            "listMessages requires either chatId, or both teamId and channelId",
          );
        }
        const messages = await graphList(context.globalArgs, url, {
          maxItems: args.maxItems,
        });
        const handle = await context.writeResource(
          "messageList",
          slugify(source, "messages"),
          {
            source,
            messages,
            count: messages.length,
            capturedAt: new Date().toISOString(),
          },
        );
        context.logger.info("Found {n} message(s) in {s}", {
          n: messages.length,
          s: source,
        });
        return { dataHandles: [handle] };
      },
    },
  },
};
