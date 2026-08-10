# @dougschaefer/ms-graph

A broad Microsoft Graph extension. One extension, nine
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
| `@dougschaefer/ms-graph-sharepoint` | `getSite`, `listFolder`, `searchDriveItems`, `downloadDriveItem` | `Sites.Read.All`, or `Sites.Selected` + a per-site grant |
| `@dougschaefer/ms-graph-intune` | `listDevices`, `getDevice`, `listCompliancePolicies`, `listConfigurationProfiles`, `listHealthScripts`, `runDeviceAction`, `exportReport` | `DeviceManagementManagedDevices.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementScripts.Read.All` (+ `DeviceManagementManagedDevices.PrivilegedOperations.All` for retire/wipe) |

All nine types share `_client.ts`'s app-only client: client-credentials token
acquisition with an in-memory token cache, a low-level `graphRequest`, a
`graphList` paging helper that auto-follows `@odata.nextLink`, and a
`graphBytes` downloader for file content. Plain `fetch` only — no native-addon
npm deps.

The `sharepoint` model provides the building blocks for document-retrieval
workflows over any library layout: `getSite` resolves the configured site,
`listFolder` snapshots a folder's children, `searchDriveItems` finds items by
name, and `downloadDriveItem` persists a file's bytes as a file artifact.

**Search behaviour depends on which permission you grant.** Graph's
`/drive/root/search(q=)` searches document *content* as well as names, but it
does not support `Sites.Selected` — under that permission it fails with HTTP 500
"General exception while processing" rather than a permission error. So
`searchDriveItems` defaults to `strategy: auto`: it tries Graph search and, on
failure, falls back to a bounded folder walk that matches names only, logging
the fallback. Set `strategy: graph` to fail loudly instead, or `walk` to skip
Graph search entirely.

The walk issues one request per folder and is slow on deep libraries — a
site-wide walk exceeded two minutes on a real library during testing. Pass
`scopePath` to bound it.

**Do not grant `Sites.ReadWrite.All` for this model.** Every method is a read,
so write access adds blast radius without adding capability. The real choice is
between the two read permissions:

`Sites.Read.All` is tenant-wide read of every site. It is the simpler grant and
the one that enables Graph's content search.

`Sites.Selected` is far narrower but comes in two halves that must both be done,
and neither works alone: consenting the app role conveys access to no site at
all, and a per-site permission entry without the consented role is never
consulted. In between you get 403 while the portal shows the permission as
granted. The payoff is that the app reaches exactly the libraries an admin has
named. Note that a Global Administrator cannot necessarily perform the per-site
half — that call runs against the *user's* SharePoint rights, not their
directory role.

This model previously read the delegated token of the active `az login` session.
That tied every call to whoever happened to be signed in and made it unrunnable
from a server, so it moved to app-only like its siblings.

## The intune model

Intune is the one Microsoft surface with no command-line alternative worth
splitting to. There is no `az intune`, the standalone Intune PowerShell SDK has
been unmaintained since 2019 and lost its global client-ID auth path in April
2024, and the current `Microsoft.Graph` PowerShell cmdlets are generated from the
same Graph description this model calls. A shell-out would buy a process spawn
and a module-version pin for an identical HTTP request, so this model is
single-transport by design.

Reads cover enrolled devices, compliance policies, configuration profiles, and
remediation health scripts. Two design points are worth knowing before you wire
them into a workflow:

**Device actions preview by default.** `runDeviceAction` fans out over a list of
device ids in one call rather than one run per device, and classifies each action
by blast radius. `sync` and `locate` dispatch immediately; `reboot`, `shutDown`,
`remoteLock`, `retire` and `wipe` return a preview naming exactly which devices
would be hit and dispatch nothing until you pass `confirm=true`. A per-device
failure is recorded and the batch continues, so one unreachable device does not
strand the rest. Be aware that Multi Admin Approval now intercepts
application-authenticated writes against protected resources, so a confirmed
action can still return 403 pending an approver — that is the tenant policy
working, not a credential problem.

**Bulk reporting goes through export jobs, not device walks.** `exportReport`
submits to `/deviceManagement/reports/exportJobs`, polls to completion, and
records the finished job. Intune builds the whole report server-side and returns
one file, so call count stays independent of fleet size and large tenants stay
clear of throttling. The download URL it records carries a short-lived SAS token
— fetch it in the following step rather than storing it for later.

Most of the surface is Graph v1.0. Settings-catalog policies and remediation
scripts are read from beta, which carries no API contract; both methods take an
`apiVersion` argument so they can move to v1.0 as coverage lands without an
extension change.

Not in scope: Win32 app deployment. Packaging needs `IntuneWinAppUtil.exe`
(Windows-only) to produce the encrypted `.intunewin` and its `Detection.xml`, and
the upload is a chunked block-blob PUT against a Graph-issued SAS URI rather than
a Graph write. That is a genuine shell-out, unlike the rest of this model, and it
belongs behind its own decision.

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
swamp model create @dougschaefer/ms-graph-calendar ms-graph-calendar \
  --global-arg clientId='${{ vault.get(azure-graph, client_id) }}' \
  --global-arg clientSecret='${{ vault.get(azure-graph, client_secret) }}' \
  --global-arg tenantId='${{ vault.get(azure-graph, tenant_id) }}'

swamp model create @dougschaefer/ms-graph-intune intune \
  --global-arg clientId='${{ vault.get(azure-graph, client_id) }}' \
  --global-arg clientSecret='${{ vault.get(azure-graph, client_secret) }}' \
  --global-arg tenantId='${{ vault.get(azure-graph, tenant_id) }}'
# ...repeat for ms-graph-users, ms-graph-places, ms-graph-groups, ms-graph-mail,
#    ms-graph-teams, ms-graph-presence
```

`sharepoint` takes the same three credentials plus a fourth argument naming the
library it reads:

```
swamp model create @dougschaefer/ms-graph-sharepoint sharepoint \
  --global-arg clientId='${{ vault.get(azure-graph, client_id) }}' \
  --global-arg clientSecret='${{ vault.get(azure-graph, client_secret) }}' \
  --global-arg tenantId='${{ vault.get(azure-graph, tenant_id) }}' \
  --global-arg siteHostPath='contoso.sharepoint.com:/sites/Clients'
```

Note the argument order: `swamp model create <type> <name>`, and global arguments
use the schema's camelCase names (`clientId`, not `client_id`) even though the
vault keys they read are snake_case.

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
| `DeviceManagementManagedDevices.Read.All` | `ms-graph-intune` (`listDevices`, `getDevice`, `exportReport`) |
| `DeviceManagementConfiguration.Read.All` | `ms-graph-intune` (`listCompliancePolicies`, `listConfigurationProfiles`) |
| `DeviceManagementScripts.Read.All` | `ms-graph-intune` (`listHealthScripts`) |
| `DeviceManagementManagedDevices.PrivilegedOperations.All` | `ms-graph-intune` (`runDeviceAction` with `retire` or `wipe`) |
| `Sites.Read.All` | `ms-graph-sharepoint` (tenant-wide read; required for content search) |
| `Sites.Selected` **+ a per-site grant** | `ms-graph-sharepoint` (narrower alternative; name-only search) |

Grant in **Entra admin center → App registrations → your app → API permissions →
Add a permission → Microsoft Graph → Application permissions**, then **Grant admin
consent**. Application `Mail.Read` is tenant-wide; scope it with an Exchange Online
`ApplicationAccessPolicy` before production use.

`Sites.Selected` is the one entry in this table that is not sufficient on its
own: consenting it grants access to nothing until an admin also adds the app to
a specific site's permissions. Calls return 403 in between, with the portal
showing the permission as granted — a confusing state worth recognising.

Intune's privileged scope is worth granting deliberately rather than by default:
without it the read methods and the sync/locate/reboot actions all still work, and
only `retire` and `wipe` fail. Grant it to a separate app registration if device
destruction should not ride the same credential as fleet inventory.

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
# Argument order is <model> <method>, not the reverse.
swamp model method run ms-graph-calendar getCurrentMeeting --input roomEmail=conf-b@example.com
swamp model method run ms-graph-users get --input idOrUpn=<entra-object-id>
swamp model method run ms-graph-places listRooms --input top=100
swamp model method run intune listDevices --input maxItems=5
```
