import { z } from "npm:zod@4.3.6";
import {
  type DataHandle,
  GRAPH_BASE,
  graphBytes,
  graphList,
  graphRequest,
  MsGraphGlobalArgsSchema,
  slugify,
  str,
} from "./_client.ts";

const SharePointGlobalArgsSchema = MsGraphGlobalArgsSchema.extend({
  siteHostPath: z
    .string()
    .describe(
      "Graph site host:path locator, e.g. contoso.sharepoint.com:/sites/Clients",
    ),
});

/** Resolved connection settings for one SharePoint document library. */
type SharePointGlobalArgs = z.infer<typeof SharePointGlobalArgsSchema>;

/** Streams bytes into a file artifact, resolving to the artifact's handle. */
interface FileWriter {
  writeAll: (bytes: Uint8Array) => Promise<DataHandle>;
}

/**
 * The swamp method context this model uses. It differs from the shared
 * MethodContext in `_client.ts` on two counts: its globalArgs carry a site
 * locator alongside the app-only credentials, and it needs createFileWriter,
 * because downloadDriveItem persists raw bytes as a file artifact rather than a
 * JSON resource.
 */
interface SharePointContext {
  globalArgs: SharePointGlobalArgs;
  logger: {
    info: (msg: string, props?: Record<string, unknown>) => void;
    warning: (msg: string, props?: Record<string, unknown>) => void;
  };
  writeResource: (
    spec: string,
    instance: string,
    data: Record<string, unknown>,
  ) => Promise<DataHandle>;
  createFileWriter: (
    spec: string,
    instance: string,
    opts: { contentType: string; tags?: Record<string, string> },
  ) => FileWriter;
}

const SiteSchema = z
  .object({
    id: z.string(),
    name: z.string().optional(),
    displayName: z.string().optional(),
    webUrl: z.string().optional(),
  })
  .passthrough();

const FolderListingSchema = z
  .object({
    path: z.string(),
    count: z.number(),
    items: z.array(z.record(z.string(), z.unknown())),
  })
  .passthrough();

const SearchResultSchema = z
  .object({
    query: z.string(),
    scopePath: z.string().optional(),
    matchedOn: z
      .string()
      .optional()
      .describe(
        "How hits were found: content+name via Graph search, or name via the folder walk.",
      ),
    truncated: z
      .boolean()
      .optional()
      .describe(
        "True when a walk hit its scan budget with folders unvisited, meaning count is a floor rather than a complete result.",
      ),
    count: z.number(),
    items: z.array(z.record(z.string(), z.unknown())),
  })
  .passthrough();

/** Trim a driveItem to the fields the listings persist. */
function slim(item: unknown): Record<string, unknown> {
  const i = (item ?? {}) as Record<string, unknown>;
  const parent = (i.parentReference ?? {}) as Record<string, unknown>;
  return {
    id: str(i.id),
    name: str(i.name),
    isFolder: i.folder !== undefined,
    size: i.size ?? null,
    lastModified: str(i.lastModifiedDateTime),
    webUrl: str(i.webUrl),
    parentPath: str(parent.path),
  };
}

/** Encode a drive-relative path for the :/{path}: URL form. */
function encPath(path: string): string {
  return path
    .split("/")
    .filter((seg) => seg.length > 0)
    .map(encodeURIComponent)
    .join("/");
}

async function siteId(g: SharePointGlobalArgs): Promise<string> {
  const { data } = await graphRequest(
    g,
    "GET",
    `${GRAPH_BASE}/sites/${g.siteHostPath}`,
  );
  return str((data as Record<string, unknown>)?.id);
}

/**
 * Find items by name by walking the folder tree, breadth-first.
 *
 * Graph's own `/drive/root/search(q=)` is the default path and is tried first.
 * It does not support the `Sites.Selected` permission — it returns HTTP 500
 * "General exception while processing" rather than a permission error — and the
 * only permission-level fix is a tenant-wide read scope such as
 * `Sites.Read.All`, which defeats the point of per-site grants. This walk is
 * what makes the method work under `Sites.Selected` anyway.
 *
 * The trade-off is real and worth stating: Graph search indexes file CONTENT,
 * while this matches on the item NAME only. Searches that relied on finding
 * words inside documents will come back empty here.
 *
 * Traversal is bounded on both ends — `maxItems` caps hits, `maxScan` caps how
 * many entries are examined — so a deep library cannot turn one call into an
 * unbounded crawl.
 */
async function walkForName(
  g: SharePointGlobalArgs,
  sid: string,
  rootPath: string,
  needle: string,
  maxItems: number,
  maxScan: number,
): Promise<{ hits: Record<string, unknown>[]; truncated: boolean }> {
  const hits: Record<string, unknown>[] = [];
  const queue: string[] = [rootPath];
  const target = needle.toLowerCase();
  let scanned = 0;
  while (queue.length > 0 && hits.length < maxItems && scanned < maxScan) {
    const cur = queue.shift() as string;
    const children = await listChildren(g, sid, cur, 999);
    for (const child of children) {
      if (scanned >= maxScan) break;
      scanned++;
      const name = str(child.name);
      if (name.toLowerCase().includes(target)) {
        hits.push(child);
        if (hits.length >= maxItems) break;
      }
      if (child.folder !== undefined) {
        queue.push(cur ? `${cur}/${name}` : name);
      }
    }
  }
  // Give up early and say so. A walk that exhausts its scan budget with folders
  // still queued looks exactly like a walk that searched everything and found
  // nothing, and silently reporting zero hits for the second reason is how a
  // caller concludes a document does not exist when it was simply never reached.
  return { hits, truncated: scanned >= maxScan && queue.length > 0 };
}

async function listChildren(
  g: SharePointGlobalArgs,
  sid: string,
  path: string,
  maxItems = 999,
): Promise<Record<string, unknown>[]> {
  const base = path
    ? `${GRAPH_BASE}/sites/${sid}/drive/root:/${encPath(path)}:/children`
    : `${GRAPH_BASE}/sites/${sid}/drive/root/children`;
  const items = await graphList(
    g,
    `${base}?$select=id,name,folder,file,size,lastModifiedDateTime,webUrl,parentReference&$top=999`,
    { maxItems },
  );
  return items as Record<string, unknown>[];
}

/**
 * `@dougschaefer/ms-graph-sharepoint` model — read-only navigation, search,
 * and file retrieval for a SharePoint document library over Microsoft Graph
 * v1.0. getSite resolves and persists the configured site; listFolder
 * snapshots one folder's children; searchDriveItems runs a Graph drive
 * search, optionally scoped to a folder path; downloadDriveItem persists one
 * file's bytes as a file artifact.
 *
 * Authenticates app-only through the shared client, like the rest of the
 * `@dougschaefer/ms-graph-*` family. It previously read the DELEGATED token of
 * the active `az login` session, which tied every call to whoever happened to
 * be signed in and made the model unrunnable from a server; app-only lets the
 * same workflows run unattended and for anyone.
 *
 * Every method here is a read. Two permission models work, and the choice is a
 * real trade-off rather than a default:
 *
 * `Sites.Read.All` is tenant-wide read of every site, and is what Graph's
 * content search requires — searchDriveItems can then match words inside
 * documents, not just filenames.
 *
 * `Sites.Selected` plus an explicit per-site grant is far narrower: the app
 * reaches exactly the libraries an admin has named and nothing else. The cost
 * is that Graph's search endpoint is unsupported under it, so searchDriveItems
 * falls back to a name-only folder walk.
 *
 * Neither needs write access. Do not grant `Sites.ReadWrite.All` for this
 * model — nothing here writes, so it would add blast radius without adding
 * capability.
 *
 *   client_id:     ${{ vault.get(azure-graph, client_id) }}
 *   client_secret: ${{ vault.get(azure-graph, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-graph, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-sharepoint",
  version: "2026.08.10.2",
  globalArguments: SharePointGlobalArgsSchema,
  resources: {
    site: {
      description: "SharePoint site metadata",
      schema: SiteSchema,
      lifetime: "infinite",
      garbageCollection: 5,
    },
    folderListing: {
      description: "One folder's children (a navigation snapshot)",
      schema: FolderListingSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    searchResult: {
      description: "Drive search hits for a query",
      schema: SearchResultSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
  },
  files: {
    document: {
      description: "A downloaded SharePoint file's bytes",
      contentType: "application/octet-stream",
      lifetime: "infinite",
      garbageCollection: 10,
    },
  },
  methods: {
    getSite: {
      description: "Resolve and persist the configured SharePoint site.",
      arguments: z.object({}),
      execute: async (
        _args: Record<string, never>,
        context: SharePointContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const g = context.globalArgs;
        const { data } = await graphRequest(
          g,
          "GET",
          `${GRAPH_BASE}/sites/${g.siteHostPath}`,
        );
        const site = (data ?? {}) as Record<string, unknown>;
        context.logger.info("Resolved site {name}", {
          name: str(site.displayName),
        });
        const handle = await context.writeResource(
          "site",
          slugify(str(site.name) || "site"),
          site as Record<string, unknown>,
        );
        return { dataHandles: [handle] };
      },
    },

    listFolder: {
      description:
        "List one folder's children (drive-relative path; empty = library root). Persists a single folderListing snapshot.",
      arguments: z.object({
        path: z
          .string()
          .default("")
          .describe(
            "Drive-relative folder path, e.g. Region/A/Acme Corporation",
          ),
        maxItems: z
          .number()
          .default(999)
          .describe("Cap on children returned"),
      }),
      execute: async (
        args: { path: string; maxItems: number },
        context: SharePointContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const items = await listChildren(g, sid, args.path, args.maxItems);
        context.logger.info("Listed {count} items under {path}", {
          count: items.length,
          path: args.path || "(root)",
        });
        const handle = await context.writeResource(
          "folderListing",
          slugify(args.path || "root"),
          {
            path: args.path,
            count: items.length,
            items: items.map(slim),
          },
        );
        return { dataHandles: [handle] };
      },
    },

    searchDriveItems: {
      description:
        "Find drive items matching a query, optionally scoped to a folder path. Persists one searchResult snapshot whose matchedOn field records how the hits were found. Graph's /drive/root/search(q=) searches document CONTENT as well as names and is used by default, but it is not supported under the Sites.Selected permission — there it fails with HTTP 500 rather than a permission error. On that failure this falls back to a bounded folder walk that matches NAMES only, so the method works under either permission model. (Permission: Sites.Read.All or Files.Read.All for content search; Sites.Selected + a per-site grant for the name-only walk)",
      arguments: z.object({
        query: z.string().describe(
          "Search terms, e.g. a project number. Matched against content and names by Graph search, or as a case-insensitive name substring by the walk.",
        ),
        scopePath: z
          .string()
          .optional()
          .describe(
            "Folder path to scope the search to; omit for the whole library",
          ),
        maxItems: z.number().default(50).describe("Cap on hits returned"),
        strategy: z
          .enum(["auto", "graph", "walk"])
          .default("auto")
          .describe(
            "auto tries Graph search and falls back to the walk; graph fails loudly instead of falling back; walk skips Graph search entirely.",
          ),
        maxScan: z.number().default(2000).describe(
          "Walk only — cap on entries examined. A site-wide walk issues one request per folder and is slow on deep libraries; prefer scopePath.",
        ),
      }),
      execute: async (
        args: {
          query: string;
          scopePath?: string;
          maxItems: number;
          strategy: "auto" | "graph" | "walk";
          maxScan: number;
        },
        context: SharePointContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        let items: unknown[] = [];
        let matchedOn = "content+name";
        let mode = args.strategy;

        if (mode === "auto" || mode === "graph") {
          const q = encodeURIComponent(args.query.replace(/'/g, "''"));
          const base = args.scopePath
            ? `${GRAPH_BASE}/sites/${sid}/drive/root:/${
              encPath(args.scopePath)
            }:/search(q='${q}')`
            : `${GRAPH_BASE}/sites/${sid}/drive/root/search(q='${q}')`;
          try {
            items = await graphList(
              g,
              `${base}?$select=id,name,folder,file,size,lastModifiedDateTime,webUrl,parentReference&$top=200`,
              { maxItems: args.maxItems },
            );
          } catch (err) {
            if (mode === "graph") throw err;
            context.logger.warning(
              "Graph drive search failed, falling back to a name-only walk: {err}",
              { err: err instanceof Error ? err.message : str(err) },
            );
            mode = "walk";
          }
        }

        let truncated = false;
        if (mode === "walk") {
          const walked = await walkForName(
            g,
            sid,
            args.scopePath ?? "",
            args.query,
            args.maxItems,
            args.maxScan,
          );
          items = walked.hits;
          truncated = walked.truncated;
          matchedOn = "name";
          if (truncated) {
            context.logger.warning(
              "Walk stopped at the maxScan cap ({cap}) with folders left unvisited — {count} hit(s) is a FLOOR, not a complete result. Narrow with scopePath or raise maxScan.",
              { cap: args.maxScan, count: items.length },
            );
          }
        }

        // The truncation notice rides the info line, not just the warning
        // above it: swamp's default log level does not render warnings, so a
        // caller running normally would see "0 item(s)" with no hint that the
        // walk gave up early.
        context.logger.info(
          "Search '{q}' matched {count} item(s) via {how}{caveat}",
          {
            q: args.query,
            count: items.length,
            how: matchedOn,
            caveat: truncated
              ? " — TRUNCATED at the maxScan cap, this count is a floor"
              : "",
          },
        );
        if (items.length >= args.maxItems) {
          context.logger.warning(
            "Hit the maxItems cap ({cap}) — results may be truncated",
            { cap: args.maxItems },
          );
        }
        const handle = await context.writeResource(
          "searchResult",
          slugify(`${args.query}-${args.scopePath ?? "site"}`),
          {
            query: args.query,
            scopePath: args.scopePath ?? "",
            matchedOn,
            truncated,
            count: items.length,
            items: (items as Record<string, unknown>[]).map(slim),
          },
        );
        return { dataHandles: [handle] };
      },
    },

    downloadDriveItem: {
      description:
        "Download one file by drive-relative path and persist its bytes as a document file artifact.",
      arguments: z.object({
        path: z
          .string()
          .describe(
            "Drive-relative file path, e.g. Region/A/Acme Corporation/12-3456 Project/signed-quote.pdf",
          ),
        resultName: z
          .string()
          .optional()
          .describe("Instance label for the file artifact"),
      }),
      execute: async (
        args: { path: string; resultName?: string },
        context: SharePointContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const { bytes, contentType } = await graphBytes(
          g,
          `${GRAPH_BASE}/sites/${sid}/drive/root:/${
            encPath(args.path)
          }:/content`,
        );
        const filename = args.path.split("/").pop() ?? "file";
        context.logger.info("Downloaded {name} ({bytes} bytes)", {
          name: filename,
          bytes: bytes.length,
        });
        const writer = context.createFileWriter(
          "document",
          slugify(args.resultName ?? filename),
          { contentType, tags: { path: args.path, filename } },
        );
        const handle = await writer.writeAll(bytes);
        return { dataHandles: [handle] };
      },
    },
  },
};
