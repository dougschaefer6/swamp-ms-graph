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

/** Standard set of message properties selected on a mailbox list read. */
const MESSAGE_SELECT =
  "id,subject,from,toRecipients,receivedDateTime,sentDateTime," +
  "isRead,hasAttachments,importance,bodyPreview,webLink";

const MessageSchema = z
  .object({
    id: z.string(),
    subject: z.string().nullish(),
    from: z.unknown().nullish(),
    toRecipients: z.array(z.unknown()).optional(),
    receivedDateTime: z.string().nullish(),
    sentDateTime: z.string().nullish(),
    isRead: z.boolean().optional(),
    hasAttachments: z.boolean().optional(),
    importance: z.string().nullish(),
    bodyPreview: z.string().nullish(),
    webLink: z.string().nullish(),
  })
  .passthrough();

/**
 * `@dougschaefer/ms-graph-mail` model — Outlook mailbox message reads over
 * Microsoft Graph v1.0 with app-only client credentials.
 *
 * Part of the broad `@dougschaefer/ms-graph-*` family. `listMessages` returns a
 * mailbox's messages (newest first, metadata + bodyPreview); `getMessage` fetches
 * one message by id including its full body.
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permission: `Mail.Read`. As of
 * this writing the Swamp Entra app has NOT been consented `Mail.Read`, so these
 * methods are built but will return HTTP 403 until an admin grants and consents
 * the scope. See the README "Scopes to grant" section.
 *
 * Note: application-permission `Mail.Read` grants tenant-wide mailbox access;
 * scope it down with an Exchange Online ApplicationAccessPolicy in production.
 *
 *   client_id:     ${{ vault.get(azure-asei, client_id) }}
 *   client_secret: ${{ vault.get(azure-asei, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-asei, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-mail",
  version: "2026.06.29.1",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    messageList: {
      description:
        "Snapshot of a mailbox's messages from GET /users/{userId}/messages, newest first, with metadata and bodyPreview.",
      schema: z.object({
        userId: z.string(),
        messages: z.array(z.unknown()),
        count: z.number(),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "5m",
      garbageCollection: 20,
    },
    message: {
      description:
        "A single Outlook message from GET /users/{userId}/messages/{id}, including the full body.",
      schema: MessageSchema,
      lifetime: "5m",
      garbageCollection: 20,
    },
  },
  methods: {
    listMessages: {
      description:
        "GET /users/{userId}/messages — list a mailbox's messages, ordered newest first, with metadata and bodyPreview. Optionally narrow with an OData $filter and cap with maxItems. Auto-follows @odata.nextLink. (Permission: Mail.Read)",
      arguments: z.object({
        userId: z.string().describe(
          "Mailbox userPrincipalName or object id, e.g. user@example.com",
        ),
        filter: z.string().optional().describe(
          'OData $filter, e.g. "isRead eq false" or "importance eq \'high\'"',
        ),
        maxItems: z.number().int().default(50).describe(
          "Hard cap on total messages collected across pages (default 50).",
        ),
      }),
      execute: async (
        args: { userId: string; filter?: string; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        let url =
          `${GRAPH_BASE}/users/${encodeURIComponent(args.userId)}/messages` +
          `?$select=${MESSAGE_SELECT}&$orderby=receivedDateTime desc&$top=50`;
        if (args.filter) url += `&$filter=${encodeURIComponent(args.filter)}`;
        const messages = await graphList(context.globalArgs, url, {
          maxItems: args.maxItems,
        });
        const handle = await context.writeResource(
          "messageList",
          slugify(args.userId, "mailbox"),
          {
            userId: args.userId,
            messages,
            count: messages.length,
            capturedAt: new Date().toISOString(),
          },
        );
        context.logger.info("Found {n} message(s) in {mbx}", {
          n: messages.length,
          mbx: args.userId,
        });
        return { dataHandles: [handle] };
      },
    },

    getMessage: {
      description:
        "GET /users/{userId}/messages/{id} — fetch one message including its full body. (Permission: Mail.Read)",
      arguments: z.object({
        userId: z.string().describe(
          "Mailbox userPrincipalName or object id that owns the message",
        ),
        messageId: z.string().describe("Graph message id"),
      }),
      execute: async (
        args: { userId: string; messageId: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("getMessage {id} for {mbx}", {
          id: args.messageId,
          mbx: args.userId,
        });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${encodeURIComponent(args.userId)}/messages/${
            encodeURIComponent(args.messageId)
          }`,
        );
        const m = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "message",
          `${slugify(args.userId, "mailbox")}-${slugify(str(m.id), "msg")}`,
          m,
        );
        return { dataHandles: [handle] };
      },
    },
  },
};
