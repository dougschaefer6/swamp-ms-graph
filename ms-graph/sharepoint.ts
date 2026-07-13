import { z } from "npm:zod@4.3.6";
import {
  azGraphAllPages,
  azGraphBytes,
  azGraphJson,
  GRAPH_BASE,
  slugify,
  str,
} from "./_client.ts";

const SharePointGlobalArgsSchema = z.object({
  siteHostPath: z
    .string()
    .describe(
      "Graph site host:path locator, e.g. contoso.sharepoint.com:/sites/Clients",
    ),
});

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

async function siteId(g: { siteHostPath: string }): Promise<string> {
  const site = await azGraphJson(`${GRAPH_BASE}/sites/${g.siteHostPath}`);
  return str(site.id);
}

async function listChildren(
  sid: string,
  path: string,
  maxItems = 999,
): Promise<Record<string, unknown>[]> {
  const base = path
    ? `${GRAPH_BASE}/sites/${sid}/drive/root:/${encPath(path)}:/children`
    : `${GRAPH_BASE}/sites/${sid}/drive/root/children`;
  const items = await azGraphAllPages(
    `${base}?$select=id,name,folder,file,size,lastModifiedDateTime,webUrl,parentReference&$top=999`,
    maxItems,
  );
  return items as Record<string, unknown>[];
}

/**
 * `@dougschaefer/ms-graph-sharepoint` model — read-only navigation, search,
 * and file retrieval for a SharePoint document library over Microsoft Graph
 * v1.0. getSite resolves and persists the configured site; listFolder
 * snapshots one folder's children; searchDriveItems runs a Graph drive
 * search, optionally scoped to a folder path; downloadDriveItem persists one
 * file's bytes as a file artifact. Unlike its ms-graph siblings this model
 * authenticates with the DELEGATED token of the active `az login` session,
 * not app-only credentials, so document access rides the signed-in operator's
 * own identity and SharePoint permissions and requires no vault
 * configuration. If it ever needs to run headless, grant an app registration
 * Sites.Selected on the specific site instead of widening to Sites.Read.All.
 */
export const model = {
  type: "@dougschaefer/ms-graph-sharepoint",
  version: "2026.07.13.3",
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
      execute: async (_args, context) => {
        const g = context.globalArgs;
        const site = await azGraphJson(
          `${GRAPH_BASE}/sites/${g.siteHostPath}`,
        );
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
            "Drive-relative folder path, e.g. Indianapolis/A/Arcwood Environmental",
          ),
        maxItems: z
          .number()
          .default(999)
          .describe("Cap on children returned"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const items = await listChildren(sid, args.path, args.maxItems);
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
        "Run a Graph drive search for a query, optionally scoped to a folder path. Persists one searchResult snapshot.",
      arguments: z.object({
        query: z.string().describe("Search terms, e.g. a project number"),
        scopePath: z
          .string()
          .optional()
          .describe("Folder path to scope the search to; omit for site-wide"),
        maxItems: z.number().default(50).describe("Cap on hits returned"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const q = encodeURIComponent(args.query.replace(/'/g, "''"));
        const base = args.scopePath
          ? `${GRAPH_BASE}/sites/${sid}/drive/root:/${
            encPath(args.scopePath)
          }:/search(q='${q}')`
          : `${GRAPH_BASE}/sites/${sid}/drive/root/search(q='${q}')`;
        const items = await azGraphAllPages(
          `${base}?$select=id,name,folder,file,size,lastModifiedDateTime,webUrl,parentReference&$top=200`,
          args.maxItems,
        );
        context.logger.info("Search '{q}' returned {count} items", {
          q: args.query,
          count: items.length,
        });
        const handle = await context.writeResource(
          "searchResult",
          slugify(`${args.query}-${args.scopePath ?? "site"}`),
          {
            query: args.query,
            scopePath: args.scopePath ?? "",
            count: items.length,
            items: items.map(slim),
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
            "Drive-relative file path, e.g. Indianapolis/A/Arcwood Environmental/00-6206 .../signed-quote.pdf",
          ),
        resultName: z
          .string()
          .optional()
          .describe("Instance label for the file artifact"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const { bytes, contentType } = await azGraphBytes(
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
