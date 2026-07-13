# @dougschaefer/ms-graph

A broad Microsoft Graph v1.0 extension. One extension, eight
model types, one shared client. It supersedes the narrow
`@dougschaefer/ms-graph-calendar` and absorbs the lookup half of
`@dougschaefer/azure-ad-user` into a Graph-native `users` model.

| Type | Methods | Graph permission |
|------|---------|------------------|
| `@dougschaefer/ms-graph-calendar` | `listEvents`, `getEvent`, `getCurrentMeeting`, `getNextMeeting` | `Calendars.Read` |
| `@dougschaefer/ms-graph-places` | `listRooms`, `getRoom` | `Place.Read.All` |
| `@dougschaefer/ms-graph-users` | `get`, `list`, `getManager`, `memberOf` | `User.Read.All` |
| `@dougschaefer/ms-graph-groups` | `list`, `get`, `listMembers` | `Group.Read.All` |
| `@dougschaefer/ms-graph-mail` | `listMessages`, `getMessage` | `Mail.Read` |
| `@dougschaefer/ms-graph-teams` | `listChats`, `listJoinedTeams`, `listChannels`, `listMessages` | `Chat.Read.All`, `ChannelMessage.Read.All` (+ `Team.ReadBasic.All`, `Channel.ReadBasic.All`) |
| `@dougschaefer/ms-graph-presence` | `getPresence` | `Presence.Read.All` |
| `@dougschaefer/ms-graph-sharepoint` | `getSite`, `listFolder`, `searchDriveItems`, `downloadDriveItem` | delegated az-session token (operator's own SharePoint permissions) |

The first seven types share `_client.ts`'s app-only client: client-credentials
token acquisition with an in-memory token cache, a low-level `graphRequest`, and
a `graphList` paging helper that auto-follows `@odata.nextLink`. Plain `fetch`
only — no native-addon npm deps.

The `sharepoint` model deliberately breaks from the app-only pattern: it reads a
document library with the DELEGATED token of the active `az login` session, so
file access rides the signed-in operator's own identity and SharePoint
permissions, requires no vault configuration and no app permission grants, and
shows the human (not an app) in SharePoint audit logs. `getSite` resolves the
configured site, `listFolder` snapshots a folder's children, `searchDriveItems`
runs a Graph drive search (site-wide or scoped to a folder path), and
`downloadDriveItem` persists a file's bytes as a file artifact — the building
blocks for document-retrieval workflows over any library layout. If it ever
needs to run headless, grant an app registration `Sites.Selected` on the
specific site instead of widening to `Sites.Read.All`.

## Authentication

App-only client-credentials flow against Microsoft Graph v1.0. Credentials come
from the existing `azure-graph` vault and are passed in as the model's
globalArguments — extensions have no vault API in model context, so the model
instance definition (or the workflow) supplies them via CEL:

```
swamp vault create azure-graph local_encryption   # skip if it already exists
swamp vault add-secret azure-graph client_id     <app-registration-client-id>
swamp vault add-secret azure-graph client_secret <app-registration-client-secret>
swamp vault add-secret azure-graph tenant_id     <entra-tenant-id>
```

## Instance creation

Each type is a separate instance; all three credential settings are identical:

```
swamp model create ms-graph-calendar \
  --type @dougschaefer/ms-graph-calendar \
  --set client_id='${{ vault.get(azure-graph, client_id) }}' \
  --set client_secret='${{ vault.get(azure-graph, client_secret) }}' \
  --set tenant_id='${{ vault.get(azure-graph, tenant_id) }}'

swamp model create ms-graph-users \
  --type @dougschaefer/ms-graph-users \
  --set client_id='${{ vault.get(azure-graph, client_id) }}' \
  --set client_secret='${{ vault.get(azure-graph, client_secret) }}' \
  --set tenant_id='${{ vault.get(azure-graph, tenant_id) }}'
# ...repeat for ms-graph-places, ms-graph-groups, ms-graph-mail, ms-graph-teams, ms-graph-presence
```

## Scopes to grant

Each model requires its Graph **application** permission (from the table at the
top) admin-consented on your app registration. Models whose permission is not
granted return HTTP 403 — grant only what you use:

| Grant to enable | Model(s) |
|-----------------|----------|
| `User.Read.All` | `ms-graph-users` |
| `Calendars.Read` | `ms-graph-calendar` |
| `Place.Read.All` | `ms-graph-places` |
| `Group.Read.All` | `ms-graph-groups` |
| `Mail.Read` | `ms-graph-mail` |
| `Chat.Read.All` | `ms-graph-teams` (chats + chat messages) |
| `ChannelMessage.Read.All` | `ms-graph-teams` (channel messages) |
| `Team.ReadBasic.All` | `ms-graph-teams` (`listJoinedTeams`) |
| `Channel.ReadBasic.All` | `ms-graph-teams` (`listChannels`) |
| `Presence.Read.All` | `ms-graph-presence` |

Grant in **Entra admin center → App registrations → your app → API permissions →
Add a permission → Microsoft Graph → Application permissions**, then **Grant admin
consent**. Application `Mail.Read` is tenant-wide; scope it with an Exchange Online
`ApplicationAccessPolicy` before production use. The `sharepoint` model needs no
app permission at all — it uses the operator's delegated az-session token.

## Migration: ms-graph-calendar and azure-ad-user

- **`ms-graph-calendar` (standalone)** is superseded by the `ms-graph-calendar`
  type in this extension. The calendar logic is ported verbatim; method names are
  `listEvents` / `getEvent` / `getCurrentMeeting` / `getNextMeeting`. (The old
  standalone `getCalendarView` is now `listEvents`; `getCurrentMeeting` /
  `getNextMeeting` are unchanged.)
- **`azure-ad-user`** (the `ad_user.ts` type inside `@dougschaefer/azure`) is a
  CLI-backed model requiring an interactive `az login`. Its read path is absorbed
  by `ms-graph-users` (`get` / `list` / `memberOf`), which uses app-only Graph REST
  and runs unattended in workflows.

  **`azure-ad-user` is retained, not retired.** Its `provision` write/create path
  (which mints a single-use temp password) was **not** ported, and it is still the
  delegate for the tracked `@dougschaefer/provision-entra-user` workflow (via the
  `entra-users` instance). Removing the type would break that live workflow, so it
  stays in `@dougschaefer/azure`. Migrate read consumers to `ms-graph-users` over
  time; revisit retiring `azure-ad-user` only once `provision` has a Graph-native
  replacement and `provision-entra-user` is rewired.

## IARS usage pattern

The meeting-agent calls `ms-graph-calendar.getCurrentMeeting` on the room mailbox
(discovered via `ms-graph-places.listRooms`). The `iars-correlate` workflow then
reverse-looks-up an asserted identity with `ms-graph-users.get` (object id → Entra
profile) and maps `department`/`jobTitle` to the AV scene.

```
# Quick live test (live-functional scopes only)
swamp model method run getCurrentMeeting ms-graph-calendar --input roomEmail=conf-b@example.com
swamp model method run get ms-graph-users --input idOrUpn=<entra-object-id>
swamp model method run listRooms ms-graph-places --input top=100
```
