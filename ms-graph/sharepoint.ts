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

const CustomerMatchSchema = z
  .object({
    customer: z.string(),
    matchCount: z.number(),
    matches: z.array(z.record(z.string(), z.unknown())),
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

const DocSetSchema = z
  .object({
    customer: z.string(),
    folderPath: z.string().optional(),
    fileCount: z.number(),
    files: z.array(z.record(z.string(), z.unknown())),
  })
  .passthrough();

/** Normalize a name for fuzzy folder matching: lowercase alphanumerics only. */
function norm(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]/g, "");
}

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
 * Find candidate customer folders in the Clients library. The library is
 * organized office → letter bucket (0, A–Z) → customer folder, with a few
 * legacy customer folders sitting directly at the drive root. Matching is
 * fuzzy in both directions on normalized names so "Arcwood Environmental"
 * finds "Arcwood Environmental" and "Mercy" finds "Mercy Health".
 */
async function findCustomerFolders(
  sid: string,
  customer: string,
  logger: {
    info: (msg: string, props?: Record<string, unknown>) => void;
  },
): Promise<Record<string, unknown>[]> {
  const want = norm(customer);
  const matches: Record<string, unknown>[] = [];
  const rootItems = await listChildren(sid, "");
  const offices: string[] = [];
  for (const item of rootItems) {
    if (item.folder === undefined) continue;
    const name = str(item.name);
    const n = norm(name);
    if (n && (n.includes(want) || want.includes(n))) {
      // Legacy customer folder at root.
      matches.push({ office: "(root)", path: name, ...slim(item) });
      continue;
    }
    offices.push(name);
  }
  const bucket = want.charAt(0).toUpperCase();
  const bucketNames = /[0-9]/.test(bucket) ? ["0"] : [bucket];
  for (const office of offices) {
    const officeItems = await listChildren(sid, office);
    const letterBuckets = officeItems.filter(
      (i) => i.folder !== undefined && str(i.name).length === 1,
    );
    let candidates: Record<string, unknown>[];
    let prefix: string;
    if (letterBuckets.length >= 5) {
      // Letter-bucketed office: descend only the customer's bucket.
      const hit = letterBuckets.find((b) =>
        bucketNames.includes(str(b.name).toUpperCase())
      );
      if (!hit) continue;
      prefix = `${office}/${str(hit.name)}`;
      candidates = await listChildren(sid, prefix);
    } else {
      prefix = office;
      candidates = officeItems;
    }
    for (const c of candidates) {
      if (c.folder === undefined) continue;
      const n = norm(str(c.name));
      if (n && (n.includes(want) || want.includes(n))) {
        matches.push({
          office,
          path: `${prefix}/${str(c.name)}`,
          ...slim(c),
        });
      }
    }
  }
  logger.info("Found {count} folder matches for {customer}", {
    count: matches.length,
    customer,
  });
  return matches;
}

/**
 * `@dougschaefer/ms-graph-sharepoint` model — read-only search and retrieval
 * over a SharePoint document library organized as a clients archive, where
 * each region/office keeps per-customer project documentation (site → office
 * → letter bucket → customer → "<project#> <title>" folders, project numbers
 * matching the ERP of record; letter buckets are auto-detected, so offices
 * that list customers directly work too). getSite/listFolder navigate;
 * findCustomerFolder resolves a
 * customer name to its folder(s) across all offices; searchDriveItems runs a
 * Graph drive search scoped to a folder; downloadDriveItem persists one file;
 * collectCustomerDocs is the fan-out used by the CPOR remediation workflow —
 * locate the customer, walk their project folders, score files against
 * keywords and project numbers, and download the best evidence in one run.
 * Unlike its ms-graph siblings this model authenticates with the DELEGATED
 * token of the active `az login` session, not app-only credentials, so
 * customer-document access rides the operator's own identity and SharePoint
 * permissions and requires no vault configuration. If this ever needs to run
 * headless, grant the app registration Sites.Selected on the site instead.
 */
export const model = {
  type: "@dougschaefer/ms-graph-sharepoint",
  version: "2026.07.13.2",
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
    customerFolderMatch: {
      description: "Customer-name to folder-path resolution result",
      schema: CustomerMatchSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    searchResult: {
      description: "Drive search hits for a query",
      schema: SearchResultSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    docSet: {
      description: "Manifest of documents collected for a customer",
      schema: DocSetSchema,
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

    findCustomerFolder: {
      description:
        "Resolve a customer name to its folder path(s) across all office folders, handling the letter-bucket layout and legacy root-level customer folders.",
      arguments: z.object({
        customer: z.string().describe("Customer name as known in the ERP"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const matches = await findCustomerFolders(
          sid,
          args.customer,
          context.logger,
        );
        const handle = await context.writeResource(
          "customerFolderMatch",
          slugify(args.customer),
          {
            customer: args.customer,
            matchCount: matches.length,
            matches,
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

    collectCustomerDocs: {
      description:
        "Fan-out evidence collector for one customer: resolve their folder, walk project subfolders, score every file against keywords and project numbers (filename and ancestor-folder hits, document types preferred), and download the top scorers as document file artifacts plus a docSet manifest. One model lock for the whole collection.",
      arguments: z.object({
        customer: z.string().describe("Customer name as known in the ERP"),
        keywords: z
          .array(z.string())
          .default([
            "signed",
            "sow",
            "statement of work",
            "proposal",
            "purchase order",
            "po",
            "quote",
            "agreement",
            "acceptance",
            "contract",
          ])
          .describe("Filename keywords that raise a file's score"),
        projectNumbers: z
          .array(z.string())
          .default([])
          .describe(
            "ERP project numbers; files under matching project folders score highest",
          ),
        maxDocs: z
          .number()
          .default(10)
          .describe("How many top-scoring files to download"),
        maxFilesScanned: z
          .number()
          .default(600)
          .describe("Cap on files walked before scoring"),
        resultName: z
          .string()
          .optional()
          .describe("Instance label for the docSet manifest"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const sid = await siteId(g);
        const matches = await findCustomerFolders(
          sid,
          args.customer,
          context.logger,
        );
        if (matches.length === 0) {
          context.logger.warning(
            "No SharePoint folder found for {customer}",
            { customer: args.customer },
          );
          const handle = await context.writeResource(
            "docSet",
            slugify(args.resultName ?? `docs-${args.customer}`),
            {
              customer: args.customer,
              folderPath: "",
              fileCount: 0,
              files: [],
            },
          );
          return { dataHandles: [handle] };
        }
        const folder = str(matches[0].path);

        // Breadth-first walk of the customer folder, two levels of subfolders
        // deep (project folders and their immediate subfolders). Each level's
        // folder listings run concurrently so large customer trees stay well
        // inside the method cancellation deadline.
        type Entry = { path: string; item: Record<string, unknown> };
        const files: Entry[] = [];
        let level: string[] = [folder];
        for (let depth = 0; depth <= 2 && level.length > 0; depth++) {
          if (files.length >= args.maxFilesScanned) break;
          const next: string[] = [];
          const CONCURRENCY = 6;
          for (let i = 0; i < level.length; i += CONCURRENCY) {
            const batch = level.slice(i, i + CONCURRENCY);
            const listings = await Promise.all(
              batch.map((p) => listChildren(sid, p)),
            );
            for (let b = 0; b < batch.length; b++) {
              for (const c of listings[b]) {
                const childPath = `${batch[b]}/${str(c.name)}`;
                if (c.folder !== undefined) {
                  if (depth < 2) next.push(childPath);
                } else if (files.length < args.maxFilesScanned) {
                  files.push({ path: childPath, item: c });
                }
              }
            }
            if (files.length >= args.maxFilesScanned) break;
          }
          level = next;
        }

        const kw = args.keywords.map((k) => k.toLowerCase());
        const projs = args.projectNumbers.map(norm);
        const EXT_SCORE: Record<string, number> = {
          pdf: 30,
          docx: 20,
          doc: 20,
          msg: 15,
          xlsx: 10,
        };
        const scored = files.map((f) => {
          const nameL = str(f.item.name).toLowerCase();
          const pathN = norm(f.path);
          let score = 0;
          for (const p of projs) {
            if (p && pathN.includes(p)) score += 100;
          }
          for (const k of kw) {
            if (nameL.includes(k)) score += 25;
          }
          const ext = nameL.split(".").pop() ?? "";
          score += EXT_SCORE[ext] ?? 0;
          if (nameL.includes("signed") || nameL.includes("executed")) {
            score += 40;
          }
          return { ...f, score };
        });
        scored.sort((a, b) =>
          b.score - a.score ||
          str(b.item.lastModifiedDateTime).localeCompare(
            str(a.item.lastModifiedDateTime),
          )
        );
        const picks = scored.filter((s) => s.score > 0).slice(0, args.maxDocs);

        const manifestFiles: Record<string, unknown>[] = [];
        const handles = [];
        for (let i = 0; i < picks.length; i++) {
          const pick = picks[i];
          const filename = str(pick.item.name);
          // The index prefix keeps instance names unique when two files
          // slugify identically (e.g. proposal.pdf next to proposal.docx).
          const savedAs = slugify(
            `sp${i + 1}-${args.customer}-${filename}`.slice(0, 60),
          );
          const { bytes, contentType } = await azGraphBytes(
            `${GRAPH_BASE}/sites/${sid}/drive/root:/${
              encPath(pick.path)
            }:/content`,
          );
          const writer = context.createFileWriter("document", savedAs, {
            contentType,
            tags: { path: pick.path, filename },
          });
          handles.push(await writer.writeAll(bytes));
          manifestFiles.push({
            name: filename,
            path: pick.path,
            savedAs,
            score: pick.score,
            size: pick.item.size ?? null,
            lastModified: str(pick.item.lastModifiedDateTime),
            webUrl: str(pick.item.webUrl),
          });
          context.logger.info("Collected {name} (score {score})", {
            name: filename,
            score: pick.score,
          });
        }

        context.logger.info(
          "Collected {picked} of {scanned} files for {customer} from {folder}",
          {
            picked: manifestFiles.length,
            scanned: files.length,
            customer: args.customer,
            folder,
          },
        );
        const manifestHandle = await context.writeResource(
          "docSet",
          slugify(args.resultName ?? `docs-${args.customer}`),
          {
            customer: args.customer,
            folderPath: folder,
            fileCount: manifestFiles.length,
            files: manifestFiles,
          },
        );
        return { dataHandles: [manifestHandle, ...handles] };
      },
    },
  },
};
