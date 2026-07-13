import { z } from "npm:zod@4.3.6";

/**
 * Shared client, schema, and context types for the `@dougschaefer/ms-graph-*`
 * model family.
 *
 * This is the broad Microsoft Graph extension for the org catalog. Seven model
 * types (calendar, places, users, groups, mail, teams, presence) all share this
 * one client: token acquisition, in-memory token caching, a low-level Graph REST
 * request helper, and an automatic `@odata.nextLink` paging helper.
 *
 * Authentication: app-only client-credentials flow against Microsoft Graph v1.0,
 * using an Entra app registration whose application permissions are admin-consented.
 *
 * Credentials are resolved from a local vault and passed in as the model's
 * globalArguments (extensions have NO vault API in model context — the workflow or
 * the model instance definition supplies them via CEL):
 *   client_id:     ${{ vault.get(azure-asei, client_id) }}
 *   client_secret: ${{ vault.get(azure-asei, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-asei, tenant_id) }}
 *
 * Plain `fetch` is used throughout — no native-addon npm libs (those break
 * `deno bundle`). Model entry files export ONLY `model`; every helper lives here
 * so deno-doc fast-check passes on each entry .ts.
 */

/** Microsoft Graph v1.0 base URL. */
export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * GlobalArgs shared by every `@dougschaefer/ms-graph-*` model: one Entra app
 * registration targeting Microsoft Graph via app-only client credentials.
 */
export const MsGraphGlobalArgsSchema = z.object({
  clientId: z.string().describe(
    "Entra app registration client ID. Use: ${{ vault.get(azure-asei, client_id) }}",
  ),
  clientSecret: z.string().meta({ sensitive: true }).describe(
    "Entra app registration client secret. Use: ${{ vault.get(azure-asei, client_secret) }}",
  ),
  tenantId: z.string().describe(
    "Entra tenant ID. Use: ${{ vault.get(azure-asei, tenant_id) }}",
  ),
  timezone: z.string().default("Eastern Standard Time").describe(
    "IANA or Windows timezone name sent as the Prefer: outlook.timezone header on calendarView calls",
  ),
  timeoutMs: z.number().int().default(30000).describe(
    "Per-request timeout in milliseconds",
  ),
});

/** Resolved connection + credentials for one Graph app registration. */
export type MsGraphGlobalArgs = z.infer<typeof MsGraphGlobalArgsSchema>;

/** Minimal shape of the swamp method context this model family uses. */
export interface MethodContext {
  globalArgs: MsGraphGlobalArgs;
  logger: {
    info: (msg: string, props?: Record<string, unknown>) => void;
    warning: (msg: string, props?: Record<string, unknown>) => void;
  };
  writeResource: (
    spec: string,
    instance: string,
    data: Record<string, unknown>,
  ) => Promise<DataHandle>;
}

/** Reference returned by `writeResource`, returned from a method's execute. */
export interface DataHandle {
  name: string;
  specName: string;
}

/** Normalized calendar meeting/event returned by the calendar model methods. */
export interface CalendarEvent {
  subject: string;
  start: string;
  end: string;
  isAllDay: boolean;
  organizer: { name: string; email: string };
  attendees: Array<{ name: string; email: string; type: string }>;
  attendeeCount: number;
}

/** Coerce any value to a string, treating null/undefined as "". */
export function str(v: unknown): string {
  return v === undefined || v === null ? "" : String(v);
}

/** Map a raw Graph event object to the normalized {@link CalendarEvent} shape. */
export function mapEvent(raw: unknown): CalendarEvent {
  const e = (raw ?? {}) as Record<string, unknown>;
  const startObj = (e.start ?? {}) as Record<string, unknown>;
  const endObj = (e.end ?? {}) as Record<string, unknown>;
  const orgObj = (e.organizer ?? {}) as Record<string, unknown>;
  const orgEmail = (orgObj.emailAddress ?? {}) as Record<string, unknown>;
  const rawAttendees = Array.isArray(e.attendees) ? e.attendees : [];
  const attendees = rawAttendees.map((a: unknown) => {
    const att = (a ?? {}) as Record<string, unknown>;
    const email = (att.emailAddress ?? {}) as Record<string, unknown>;
    return {
      name: str(email.name),
      email: str(email.address),
      type: str(att.type),
    };
  });
  return {
    subject: str(e.subject),
    start: str(startObj.dateTime),
    end: str(endObj.dateTime),
    isAllDay: Boolean(e.isAllDay),
    organizer: {
      name: str(orgEmail.name),
      email: str(orgEmail.address),
    },
    attendees,
    attendeeCount: attendees.length,
  };
}

/** In-memory token cache (single-process lifetime). */
let _tokenCache: { token: string; expiresAt: number } | null = null;

/**
 * Acquire a client-credentials Bearer token for Microsoft Graph.
 * The token is cached in memory until 60 s before expiry so we don't
 * re-authenticate on every method call.
 *
 * @param g - resolved globalArgs
 */
export async function acquireToken(g: MsGraphGlobalArgs): Promise<string> {
  const now = Date.now();
  if (_tokenCache && _tokenCache.expiresAt > now) {
    return _tokenCache.token;
  }
  const url = `https://login.microsoftonline.com/${
    encodeURIComponent(g.tenantId)
  }/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: g.clientId,
    client_secret: g.clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), g.timeoutMs);
  let res: Response;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      signal: ctrl.signal,
    });
  } finally {
    clearTimeout(timer);
  }
  const text = await res.text();
  let data: Record<string, unknown> = {};
  try {
    data = JSON.parse(text);
  } catch {
    // fall through
  }
  if (!res.ok) {
    const desc = str(
      data.error_description ?? data.error ?? text.slice(0, 300),
    );
    throw new Error(`Graph token request failed HTTP ${res.status}: ${desc}`);
  }
  const token = str(data.access_token);
  const expiresIn = Number(data.expires_in ?? 3600);
  _tokenCache = { token, expiresAt: now + (expiresIn - 60) * 1000 };
  return token;
}

/** Extract Graph's `{ error: { message } }` detail from a response body. */
function graphErrorMessage(data: unknown): string | undefined {
  if (data && typeof data === "object") {
    const d = data as Record<string, unknown>;
    const err = d.error as Record<string, unknown> | undefined;
    if (err && typeof err.message === "string") return err.message;
  }
  return undefined;
}

/**
 * Low-level Microsoft Graph REST request.
 * Acquires a token, appends Accept and optional Prefer headers, and throws on
 * non-2xx, surfacing Graph's `error.message`.
 *
 * @param g - resolved globalArgs
 * @param method - HTTP verb
 * @param url - full URL (Graph v1.0 base is `https://graph.microsoft.com/v1.0`)
 * @param opts - optional Prefer header value and additional query params
 */
export async function graphRequest(
  g: MsGraphGlobalArgs,
  method: string,
  url: string,
  opts: {
    query?: Record<string, string>;
    prefer?: string;
  } = {},
): Promise<{ status: number; data: unknown }> {
  const token = await acquireToken(g);
  const parsed = new URL(url);
  if (opts.query) {
    for (const [k, v] of Object.entries(opts.query)) {
      if (v !== undefined && v !== "") parsed.searchParams.set(k, v);
    }
  }
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
  };
  if (opts.prefer) {
    headers["Prefer"] = opts.prefer;
  }
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), g.timeoutMs);
  let res: Response;
  try {
    res = await fetch(parsed.toString(), {
      method,
      headers,
      signal: ctrl.signal,
    });
  } finally {
    clearTimeout(timer);
  }
  const text = await res.text();
  let data: unknown = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = text;
    }
  }
  if (!res.ok) {
    const msg = graphErrorMessage(data) ?? text.slice(0, 400);
    throw new Error(
      `Graph ${method} ${parsed.pathname} -> HTTP ${res.status}: ${msg}`,
    );
  }
  return { status: res.status, data };
}

/**
 * Page a Graph collection endpoint, auto-following `@odata.nextLink` until the
 * collection is exhausted or `maxItems` is reached. Returns the accumulated
 * `value` array. The first `url` should already carry any `$select`/`$filter`/
 * `$top`/`$orderby` query parameters.
 *
 * @param g - resolved globalArgs
 * @param url - first page URL (Graph v1.0)
 * @param opts - optional Prefer header and a hard cap on total items collected
 */
export async function graphList(
  g: MsGraphGlobalArgs,
  url: string,
  opts: { prefer?: string; maxItems?: number } = {},
): Promise<unknown[]> {
  const out: unknown[] = [];
  const cap = opts.maxItems ?? Number.POSITIVE_INFINITY;
  let nextLink: string | null = url;
  while (nextLink) {
    const { data } = await graphRequest(g, "GET", nextLink, {
      prefer: opts.prefer,
    });
    const d = (data ?? {}) as Record<string, unknown>;
    const page = Array.isArray(d.value) ? d.value : [];
    for (const item of page) {
      out.push(item);
      if (out.length >= cap) return out;
    }
    nextLink = typeof d["@odata.nextLink"] === "string"
      ? d["@odata.nextLink"]
      : null;
  }
  return out;
}

/**
 * Turn an email address or arbitrary identifier into a short, file-name-safe
 * slug used to name the data instance a method writes.
 */
export function slugify(value: string, fallback = "item"): string {
  return value
    .replace(/@.*$/, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48) || fallback;
}

/** In-memory cache for the delegated az-session Graph token. */
let _azTokenCache: { token: string; expiresAt: number } | null = null;

/**
 * Acquire a DELEGATED Microsoft Graph token from the active `az login`
 * session (the signed-in human, not an app registration). Used by the
 * sharepoint model, where document access should ride the operator's own
 * identity and SharePoint permissions rather than an app-only grant. The
 * token arrives in az's JSON output (never on a command line) and is cached
 * until shortly before expiry.
 */
export async function azGraphToken(): Promise<string> {
  const now = Date.now();
  if (_azTokenCache && _azTokenCache.expiresAt > now) {
    return _azTokenCache.token;
  }
  const cmd = new Deno.Command("az", {
    args: [
      "account",
      "get-access-token",
      "--resource",
      "https://graph.microsoft.com",
      "--output",
      "json",
    ],
    stdout: "piped",
    stderr: "piped",
  });
  const result = await cmd.output();
  if (result.code !== 0) {
    const stderr = new TextDecoder().decode(result.stderr);
    throw new Error(`az graph token acquisition failed: ${stderr}`);
  }
  const data = JSON.parse(new TextDecoder().decode(result.stdout)) as Record<
    string,
    unknown
  >;
  const token = str(data.accessToken);
  // expiresOn is a local-time string; refresh conservatively every 20 min.
  _azTokenCache = { token, expiresAt: now + 20 * 60 * 1000 };
  return token;
}

/** GET a Graph URL with the delegated az-session token; throws on non-2xx. */
export async function azGraphJson(
  url: string,
): Promise<Record<string, unknown>> {
  const token = await azGraphToken();
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  const text = await res.text();
  let data: Record<string, unknown> = {};
  try {
    data = JSON.parse(text);
  } catch {
    // fall through
  }
  if (!res.ok) {
    const detail = graphErrorMessage(data) ?? text.slice(0, 300);
    throw new Error(`Graph GET ${url} failed HTTP ${res.status}: ${detail}`);
  }
  return data;
}

/** Follow @odata.nextLink pages for a delegated az-session GET. */
export async function azGraphAllPages(
  url: string,
  maxItems = 2000,
): Promise<unknown[]> {
  const out: unknown[] = [];
  let nextLink: string | null = url;
  while (nextLink) {
    const d = await azGraphJson(nextLink);
    const page = Array.isArray(d.value) ? d.value : [];
    for (const item of page) {
      out.push(item);
      if (out.length >= maxItems) return out;
    }
    nextLink = typeof d["@odata.nextLink"] === "string"
      ? d["@odata.nextLink"]
      : null;
  }
  return out;
}

/** Download a Graph URL's raw bytes with the delegated az-session token. */
export async function azGraphBytes(
  url: string,
): Promise<{ bytes: Uint8Array; contentType: string }> {
  const token = await azGraphToken();
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `Graph download ${url} failed HTTP ${res.status}: ${text.slice(0, 200)}`,
    );
  }
  return {
    bytes: new Uint8Array(await res.arrayBuffer()),
    contentType: res.headers.get("content-type") ?? "application/octet-stream",
  };
}
