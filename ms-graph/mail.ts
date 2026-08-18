import { z } from "npm:zod@4.3.6";
import {
  type DataHandle,
  GRAPH_BASE,
  graphBytes,
  graphList,
  graphRequest,
  type MethodContext,
  type MsGraphGlobalArgs,
  MsGraphGlobalArgsSchema,
  slugify,
  str,
} from "./_client.ts";

/** Streams bytes into a file artifact, resolving to the artifact's handle. */
interface FileWriter {
  writeAll: (bytes: Uint8Array) => Promise<DataHandle>;
}

/**
 * The method context this model uses. It extends the shared MethodContext with
 * createFileWriter, because getMimeContent persists a message's raw RFC822
 * bytes as a file artifact rather than as a JSON resource.
 */
interface MailContext extends MethodContext {
  globalArgs: MsGraphGlobalArgs;
  createFileWriter: (
    spec: string,
    instance: string,
    opts: { contentType: string; tags?: Record<string, string> },
  ) => FileWriter;
}

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
 * one message by id including its full body; `getMimeContent` persists a
 * message's unmodified RFC822 source as a file artifact, preserving the
 * Authentication-Results and DKIM-Signature headers that make an email usable as
 * evidence of who actually sent it.
 *
 * Authentication uses the shared app-only client from `_client.ts` against the
 * configured Entra app vault. Required application permission: `Mail.Read`.
 * These methods return HTTP 403 until an admin grants and consents that scope
 * on your app registration. See the README "Scopes to grant" section.
 *
 * Note: application-permission `Mail.Read` grants tenant-wide mailbox access;
 * scope it down with an Exchange Online ApplicationAccessPolicy in production.
 *
 *   client_id:     ${{ vault.get(azure-graph, client_id) }}
 *   client_secret: ${{ vault.get(azure-graph, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-graph, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-mail",
  version: "2026.08.17.1",
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
  files: {
    messageMime: {
      description:
        "One message's unmodified RFC822 source from GET /messages/{id}/$value, headers intact.",
      contentType: "message/rfc822",
      lifetime: "infinite",
      garbageCollection: 10,
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
        // Results are newest-first, so hitting the cap means the OLD end of the
        // requested window was silently dropped — a search over the remainder
        // would report "nothing found" when it simply never looked. Say so
        // loudly. This is logger.info on purpose: warning-level never renders
        // without -v, and a truncation notice nobody sees is worse than none.
        if (messages.length >= args.maxItems) {
          context.logger.info(
            "TRUNCATED at maxItems={cap} — newest {cap} only; older messages in this filter were NOT read. " +
              "Narrow the window with an upper bound (receivedDateTime lt ...) or raise maxItems before treating an empty result as a negative.",
            { cap: args.maxItems },
          );
        }
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

    getMimeContent: {
      description:
        "GET /users/{userId}/messages/{id}/$value — persist a message's unmodified RFC822 source as a messageMime file artifact. Unlike getMessage, which returns a parsed body, this preserves the full header block including Authentication-Results, DKIM-Signature and Received chain, so the file stands on its own as evidence of who sent the message and when. Write it to disk with a .eml extension. (Permission: Mail.Read)",
      arguments: z.object({
        userId: z.string().describe(
          "Mailbox userPrincipalName or object id that owns the message",
        ),
        messageId: z.string().describe("Graph message id"),
        resultName: z.string().optional().describe(
          "Instance label for the file artifact; defaults to a slug of the mailbox and message id",
        ),
      }),
      execute: async (
        args: { userId: string; messageId: string; resultName?: string },
        context: MailContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const { bytes, contentType } = await graphBytes(
          context.globalArgs,
          `${GRAPH_BASE}/users/${encodeURIComponent(args.userId)}/messages/${
            encodeURIComponent(args.messageId)
          }/$value`,
        );
        context.logger.info(
          "Fetched RFC822 source for {id} from {mbx} ({bytes} bytes)",
          {
            id: args.messageId,
            mbx: args.userId,
            bytes: bytes.length,
          },
        );
        const writer = context.createFileWriter(
          "messageMime",
          slugify(
            args.resultName ??
              `${slugify(args.userId, "mailbox")}-${
                slugify(args.messageId, "msg")
              }`,
          ),
          {
            // Graph answers $value as text/plain; the payload is RFC822 either
            // way, so label the artifact for what it actually is.
            contentType: contentType.startsWith("text/plain")
              ? "message/rfc822"
              : contentType,
            tags: { userId: args.userId, messageId: args.messageId },
          },
        );
        const handle = await writer.writeAll(bytes);
        return { dataHandles: [handle] };
      },
    },
  },
};
