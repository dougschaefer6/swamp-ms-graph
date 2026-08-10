import { z } from "npm:zod@4.3.6";
import {
  type DataHandle,
  GRAPH_BASE,
  GRAPH_BETA,
  graphList,
  graphRequest,
  type MethodContext,
  MsGraphGlobalArgsSchema,
  slugify,
  str,
} from "./_client.ts";

/**
 * Standard property set selected on every managed-device read. Chosen to answer
 * the fleet questions without a second call: who has it, what it is, whether it
 * is compliant, and when it last checked in.
 */
const DEVICE_SELECT = [
  "id",
  "deviceName",
  "managedDeviceName",
  "userPrincipalName",
  "userDisplayName",
  "operatingSystem",
  "osVersion",
  "complianceState",
  "managementState",
  "managementAgent",
  "managedDeviceOwnerType",
  "deviceEnrollmentType",
  "deviceRegistrationState",
  "enrolledDateTime",
  "lastSyncDateTime",
  "model",
  "manufacturer",
  "serialNumber",
  "azureADDeviceId",
  "isEncrypted",
  "isSupervised",
  "jailBroken",
].join(",");

/**
 * Device actions this model can dispatch, mapped to their Graph action segment
 * and a blast-radius tier. The tier drives the confirmation gate: `safe`
 * actions run without confirmation, everything else is previewed unless the
 * caller explicitly confirms.
 *
 * `retire` removes company data and unenrolls but leaves personal data intact;
 * `wipe` is a factory reset. Both are irreversible from Intune's side, which is
 * why they sit in the `destructive` tier together.
 */
const DEVICE_ACTIONS = {
  sync: { segment: "syncDevice", tier: "safe" },
  locate: { segment: "locateDevice", tier: "safe" },
  reboot: { segment: "rebootNow", tier: "disruptive" },
  shutDown: { segment: "shutDown", tier: "disruptive" },
  remoteLock: { segment: "remoteLock", tier: "disruptive" },
  retire: { segment: "retire", tier: "destructive" },
  wipe: { segment: "wipe", tier: "destructive" },
} as const;

/** Action names accepted by the device-action method. */
type DeviceActionName = keyof typeof DEVICE_ACTIONS;

const ManagedDeviceSchema = z
  .object({
    id: z.string(),
    deviceName: z.string().nullish(),
    userPrincipalName: z.string().nullish(),
    operatingSystem: z.string().nullish(),
    osVersion: z.string().nullish(),
    complianceState: z.string().nullish(),
    managementState: z.string().nullish(),
    lastSyncDateTime: z.string().nullish(),
    serialNumber: z.string().nullish(),
    model: z.string().nullish(),
    manufacturer: z.string().nullish(),
  })
  .passthrough();

const CompliancePolicySchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
    "@odata.type": z.string().nullish(),
    createdDateTime: z.string().nullish(),
    lastModifiedDateTime: z.string().nullish(),
    version: z.number().nullish(),
  })
  .passthrough();

const ConfigurationProfileSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    name: z.string().nullish(),
    description: z.string().nullish(),
    profileKind: z
      .string()
      .describe(
        "Which Intune surface this came from: deviceConfiguration (classic templates, v1.0) or configurationPolicy (settings catalog, beta).",
      ),
    "@odata.type": z.string().nullish(),
    platforms: z.string().nullish(),
    technologies: z.string().nullish(),
  })
  .passthrough();

const HealthScriptSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
    publisher: z.string().nullish(),
    runAsAccount: z.string().nullish(),
    runAs32Bit: z.boolean().nullish(),
    enforceSignatureCheck: z.boolean().nullish(),
    isGlobalScript: z.boolean().nullish(),
  })
  .passthrough();

const DeviceActionSchema = z
  .object({
    action: z.string(),
    tier: z.string(),
    dispatched: z
      .boolean()
      .describe(
        "Whether Graph was actually called. False means the run was a preview held back by the confirmation gate — no device was touched.",
      ),
    requestedCount: z.number(),
    succeeded: z.array(z.string()),
    failed: z.array(z.object({ id: z.string(), error: z.string() })),
    skipped: z
      .array(z.string())
      .describe(
        "Device ids not dispatched because the run was an unconfirmed preview.",
      ),
  })
  .passthrough();

const ReportExportSchema = z
  .object({
    id: z.string(),
    reportName: z.string().nullish(),
    status: z.string().nullish(),
    format: z.string().nullish(),
    requestDateTime: z.string().nullish(),
    expirationDateTime: z.string().nullish(),
    url: z
      .string()
      .nullish()
      .describe(
        "Download URL for the finished report. Carries a short-lived SAS token and expires — fetch it promptly in a following step rather than storing it for later.",
      ),
  })
  .passthrough();

/**
 * Dispatch one Graph device action and normalize the outcome, so a fan-out over
 * many devices records a per-device result instead of aborting the whole run on
 * the first failure.
 */
async function dispatchAction(
  context: MethodContext,
  deviceId: string,
  segment: string,
  body: unknown,
): Promise<{ ok: true } | { ok: false; error: string }> {
  const url = `${GRAPH_BASE}/deviceManagement/managedDevices/${
    encodeURIComponent(deviceId)
  }/${segment}`;
  try {
    await graphRequest(
      context.globalArgs,
      "POST",
      url,
      body === undefined ? {} : { body },
    );
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : str(err) };
  }
}

/**
 * `@dougschaefer/ms-graph-intune` model — Microsoft Intune device management
 * over Microsoft Graph with app-only client credentials.
 *
 * Part of the broad `@dougschaefer/ms-graph-*` family, sharing the token cache,
 * paging, and error handling in `_client.ts`. Intune has no CLI of its own —
 * there is no `az intune`, and the Microsoft.Graph PowerShell cmdlets are
 * generated from the same Graph description this model calls — so unlike the
 * azure extension there is no CLI transport worth splitting to. Everything here
 * is one wire.
 *
 * Reads cover enrolled devices, compliance policies, configuration profiles
 * (both classic templates and the settings catalog), and remediation health
 * scripts. Writes are limited to device actions, which are gated: anything more
 * disruptive than a sync previews by default and only dispatches when the caller
 * passes `confirm`. Bulk reporting goes through the async export-job API rather
 * than walking devices one at a time, which keeps call count independent of
 * fleet size.
 *
 * Most of the surface is Graph v1.0. The settings catalog and remediation
 * scripts are read from beta, which has no API contract — those methods take an
 * `apiVersion` argument so they can be moved to v1.0 as coverage lands without
 * an extension change.
 *
 * Required application permissions, admin-consented on the app registration:
 * `DeviceManagementManagedDevices.Read.All` for device reads,
 * `DeviceManagementConfiguration.Read.All` for compliance policies and
 * configuration profiles, `DeviceManagementScripts.Read.All` for the
 * remediation health scripts (a separate scope from configuration, verified
 * against a live tenant), and
 * `DeviceManagementManagedDevices.PrivilegedOperations.All` for the retire
 * and wipe actions. Methods return HTTP 403 until those are granted. Note that
 * Multi Admin Approval intercepts application-authenticated writes against
 * protected resources, so a confirmed device action can 403 pending approval
 * even with the scope in place.
 *
 *   client_id:     ${{ vault.get(azure-graph, client_id) }}
 *   client_secret: ${{ vault.get(azure-graph, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-graph, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-intune",
  version: "2026.08.10.2",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    managedDevice: {
      description:
        "An Intune-enrolled device from /deviceManagement/managedDevices, selected to the standard inventory fields (owner, OS, compliance state, last sync, hardware identifiers).",
      schema: ManagedDeviceSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    compliancePolicy: {
      description:
        "A device compliance policy from /deviceManagement/deviceCompliancePolicies. The @odata.type identifies the platform the policy targets.",
      schema: CompliancePolicySchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    configurationProfile: {
      description:
        "A device configuration profile — either a classic template from /deviceManagement/deviceConfigurations or a settings-catalog policy from /deviceManagement/configurationPolicies, distinguished by profileKind.",
      schema: ConfigurationProfileSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    healthScript: {
      description:
        "A remediation script pair (detection plus remediation) from /deviceManagement/deviceHealthScripts.",
      schema: HealthScriptSchema,
      lifetime: "infinite",
      garbageCollection: 10,
    },
    deviceAction: {
      description:
        "The outcome of one fan-out device action: which devices succeeded, which failed and why, and which were skipped because the run was an unconfirmed preview.",
      schema: DeviceActionSchema,
      lifetime: "infinite",
      garbageCollection: 20,
    },
    reportExport: {
      description:
        "A finished Intune report export job, including the short-lived download URL for the zipped CSV or JSON payload.",
      schema: ReportExportSchema,
      lifetime: "infinite",
      garbageCollection: 20,
    },
  },
  methods: {
    listDevices: {
      description:
        "GET /deviceManagement/managedDevices — enumerate enrolled devices, selected to the standard inventory fields. Auto-follows @odata.nextLink. Note that managedDevices supports $filter on only a subset of properties (operatingSystem, complianceState, ownerType and similar); filtering on an unsupported property returns HTTP 400 rather than an empty set. (Permission: DeviceManagementManagedDevices.Read.All)",
      arguments: z.object({
        filter: z.string().optional().describe(
          "OData $filter over supported properties, e.g. \"complianceState eq 'noncompliant'\" or \"operatingSystem eq 'Windows'\"",
        ),
        maxItems: z.number().int().default(500).describe(
          "Hard cap on total devices collected across pages (default 500).",
        ),
      }),
      execute: async (
        args: { filter?: string; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        let url =
          `${GRAPH_BASE}/deviceManagement/managedDevices?$select=${DEVICE_SELECT}&$top=100`;
        if (args.filter) url += `&$filter=${encodeURIComponent(args.filter)}`;
        const devices = await graphList(context.globalArgs, url, {
          maxItems: args.maxItems,
        });
        const noncompliant =
          devices.filter((d) =>
            str((d as Record<string, unknown>).complianceState) ===
              "noncompliant"
          ).length;
        context.logger.info(
          "Found {count} device(s), {noncompliant} noncompliant",
          { count: devices.length, noncompliant },
        );
        const handles: DataHandle[] = [];
        for (const d of devices) {
          const dev = (d ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "managedDevice",
              slugify(str(dev.id), "device"),
              dev,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    getDevice: {
      description:
        "GET /deviceManagement/managedDevices/{id} — resolve one enrolled device by its Intune managed-device id, selected to the standard inventory fields. (Permission: DeviceManagementManagedDevices.Read.All)",
      arguments: z.object({
        id: z.string().describe("Intune managed-device id (GUID)"),
      }),
      execute: async (
        args: { id: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("Resolving device {id}", { id: args.id });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/deviceManagement/managedDevices/${
            encodeURIComponent(args.id)
          }?$select=${DEVICE_SELECT}`,
        );
        const dev = (data ?? {}) as Record<string, unknown>;
        const handle = await context.writeResource(
          "managedDevice",
          slugify(str(dev.id), "device"),
          dev,
        );
        context.logger.info("Resolved {name} ({state})", {
          name: str(dev.deviceName),
          state: str(dev.complianceState),
        });
        return { dataHandles: [handle] };
      },
    },

    listCompliancePolicies: {
      description:
        "GET /deviceManagement/deviceCompliancePolicies — enumerate compliance policies across platforms, optionally expanding their group assignments. (Permission: DeviceManagementConfiguration.Read.All)",
      arguments: z.object({
        includeAssignments: z.boolean().default(false).describe(
          "Expand each policy's assignments so the groups it targets come back in the same call.",
        ),
        maxItems: z.number().int().default(200).describe(
          "Hard cap on total policies collected across pages (default 200).",
        ),
      }),
      execute: async (
        args: { includeAssignments: boolean; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        let url =
          `${GRAPH_BASE}/deviceManagement/deviceCompliancePolicies?$top=100`;
        if (args.includeAssignments) url += "&$expand=assignments";
        const policies = await graphList(context.globalArgs, url, {
          maxItems: args.maxItems,
        });
        context.logger.info("Found {count} compliance policy/policies", {
          count: policies.length,
        });
        const handles: DataHandle[] = [];
        for (const p of policies) {
          const pol = (p ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "compliancePolicy",
              slugify(str(pol.id), "policy"),
              pol,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    listConfigurationProfiles: {
      description:
        "Enumerate device configuration profiles. Always reads the classic templates from /deviceManagement/deviceConfigurations; with includeSettingsCatalog it also reads /deviceManagement/configurationPolicies, which is beta-only. Each resource carries a profileKind marking which surface it came from. (Permission: DeviceManagementConfiguration.Read.All)",
      arguments: z.object({
        includeSettingsCatalog: z.boolean().default(true).describe(
          "Also collect settings-catalog policies. These live only on the beta endpoint, so their shape is not contract-guaranteed.",
        ),
        maxItems: z.number().int().default(200).describe(
          "Hard cap on profiles collected from each surface (default 200 each).",
        ),
      }),
      execute: async (
        args: { includeSettingsCatalog: boolean; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const handles: DataHandle[] = [];
        const templates = await graphList(
          context.globalArgs,
          `${GRAPH_BASE}/deviceManagement/deviceConfigurations?$top=100`,
          { maxItems: args.maxItems },
        );
        for (const t of templates) {
          const tpl = (t ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "configurationProfile",
              slugify(str(tpl.id), "profile"),
              { ...tpl, profileKind: "deviceConfiguration" },
            ),
          );
        }
        let catalogCount = 0;
        if (args.includeSettingsCatalog) {
          const catalog = await graphList(
            context.globalArgs,
            `${GRAPH_BETA}/deviceManagement/configurationPolicies?$top=100`,
            { maxItems: args.maxItems },
          );
          catalogCount = catalog.length;
          for (const c of catalog) {
            const pol = (c ?? {}) as Record<string, unknown>;
            handles.push(
              await context.writeResource(
                "configurationProfile",
                slugify(str(pol.id), "profile"),
                { ...pol, profileKind: "configurationPolicy" },
              ),
            );
          }
        }
        context.logger.info(
          "Found {templates} template profile(s) and {catalog} settings-catalog policy/policies",
          { templates: templates.length, catalog: catalogCount },
        );
        return { dataHandles: handles };
      },
    },

    listHealthScripts: {
      description:
        "GET /deviceManagement/deviceHealthScripts — enumerate remediation script pairs (detection plus remediation), the mechanism behind Intune Remediations. Defaults to the beta endpoint because that is where the resource is published. Note this is the one read that does NOT ride DeviceManagementConfiguration.Read.All — scripts are a separate scope. (Permission: DeviceManagementScripts.Read.All)",
      arguments: z.object({
        apiVersion: z.enum(["beta", "v1.0"]).default("beta").describe(
          "Graph version to read from. Move this to v1.0 once the resource graduates.",
        ),
        maxItems: z.number().int().default(200).describe(
          "Hard cap on total scripts collected across pages (default 200).",
        ),
      }),
      execute: async (
        args: { apiVersion: "beta" | "v1.0"; maxItems: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const base = args.apiVersion === "beta" ? GRAPH_BETA : GRAPH_BASE;
        const scripts = await graphList(
          context.globalArgs,
          `${base}/deviceManagement/deviceHealthScripts?$top=100`,
          { maxItems: args.maxItems },
        );
        context.logger.info("Found {count} remediation script(s)", {
          count: scripts.length,
        });
        const handles: DataHandle[] = [];
        for (const s of scripts) {
          const scr = (s ?? {}) as Record<string, unknown>;
          handles.push(
            await context.writeResource(
              "healthScript",
              slugify(str(scr.id), "script"),
              scr,
            ),
          );
        }
        return { dataHandles: handles };
      },
    },

    runDeviceAction: {
      description:
        "Dispatch one device action across many devices in a single call, recording a per-device result so one failure does not abort the batch. Sync and locate run immediately; reboot, shutDown, remoteLock, retire and wipe only dispatch when confirm is true, and otherwise return a preview listing exactly which devices would be acted on. Retire removes company data and unenrolls; wipe is a factory reset — both are irreversible from Intune. (Permissions: DeviceManagementManagedDevices.Read.All, plus DeviceManagementManagedDevices.PrivilegedOperations.All for retire and wipe)",
      arguments: z.object({
        deviceIds: z.array(z.string()).min(1).describe(
          "Intune managed-device ids to act on.",
        ),
        action: z
          .enum([
            "sync",
            "locate",
            "reboot",
            "shutDown",
            "remoteLock",
            "retire",
            "wipe",
          ])
          .describe("Action to dispatch to every listed device."),
        confirm: z.boolean().default(false).describe(
          "Required for anything beyond sync and locate. Left false, the method previews the batch without calling Graph.",
        ),
        keepUserData: z.boolean().default(false).describe(
          "Wipe only — preserve user data instead of a full factory reset. Ignored by every other action.",
        ),
      }),
      execute: async (
        args: {
          deviceIds: string[];
          action: DeviceActionName;
          confirm: boolean;
          keepUserData: boolean;
        },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const spec = DEVICE_ACTIONS[args.action];
        const gated = spec.tier !== "safe" && !args.confirm;
        const succeeded: string[] = [];
        const failed: Array<{ id: string; error: string }> = [];
        const skipped: string[] = [];

        if (gated) {
          // info, not warning: swamp's default log level does not render
          // warnings, and "nothing was dispatched" is the single most important
          // thing a caller of a destructive action needs to see.
          context.logger.info(
            "PREVIEW ONLY — {action} ({tier}) would run on {count} device(s). NOTHING was dispatched. Re-run with confirm=true to act.",
            {
              action: args.action,
              tier: spec.tier,
              count: args.deviceIds.length,
            },
          );
          skipped.push(...args.deviceIds);
        } else {
          const body = args.action === "wipe"
            ? { keepUserData: args.keepUserData }
            : undefined;
          for (const id of args.deviceIds) {
            const result = await dispatchAction(
              context,
              id,
              spec.segment,
              body,
            );
            if (result.ok) {
              succeeded.push(id);
            } else {
              failed.push({ id, error: result.error });
              context.logger.warning("{action} failed on {id}: {error}", {
                action: args.action,
                id,
                error: result.error,
              });
            }
          }
          context.logger.info(
            "{action} dispatched: {ok} succeeded, {bad} failed",
            { action: args.action, ok: succeeded.length, bad: failed.length },
          );
        }

        const handle = await context.writeResource(
          "deviceAction",
          `${slugify(args.action, "action")}-${Date.now()}`,
          {
            action: args.action,
            tier: spec.tier,
            dispatched: !gated,
            requestedCount: args.deviceIds.length,
            succeeded,
            failed,
            skipped,
          },
        );
        return { dataHandles: [handle] };
      },
    },

    exportReport: {
      description:
        "POST /deviceManagement/reports/exportJobs then poll to completion — ask Intune to build a whole report server-side and hand back one file, instead of walking devices individually. This keeps call count independent of fleet size and stays clear of throttling on large tenants. Writes the finished job including its download URL, which carries a short-lived SAS token; fetch it in a following step rather than storing it. (Permission: DeviceManagementManagedDevices.Read.All)",
      arguments: z.object({
        reportName: z.string().describe(
          "Intune report to build, e.g. Devices, DevicesWithInventory, DeviceCompliance, DeviceNonCompliance, or DeviceInstallStatusByApp.",
        ),
        format: z.enum(["csv", "json"]).default("csv").describe(
          "Payload format inside the returned zip.",
        ),
        filter: z.string().optional().describe(
          "Report-specific filter expression, passed through to the export job unchanged.",
        ),
        select: z.array(z.string()).optional().describe(
          "Columns to include. Omit for the report's default column set.",
        ),
        apiVersion: z.enum(["v1.0", "beta"]).default("v1.0").describe(
          "Graph version to submit against. Several report names are published only on beta.",
        ),
        timeoutSeconds: z.number().int().default(300).describe(
          "How long to poll before giving up on the job (default 300).",
        ),
      }),
      execute: async (
        args: {
          reportName: string;
          format: "csv" | "json";
          filter?: string;
          select?: string[];
          apiVersion: "v1.0" | "beta";
          timeoutSeconds: number;
        },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const base = args.apiVersion === "beta" ? GRAPH_BETA : GRAPH_BASE;
        const body: Record<string, unknown> = {
          reportName: args.reportName,
          format: args.format,
        };
        if (args.filter) body.filter = args.filter;
        if (args.select) body.select = args.select;

        context.logger.info("Submitting export job for {report}", {
          report: args.reportName,
        });
        const { data: created } = await graphRequest(
          context.globalArgs,
          "POST",
          `${base}/deviceManagement/reports/exportJobs`,
          { body },
        );
        const job = (created ?? {}) as Record<string, unknown>;
        const jobId = str(job.id);
        if (!jobId) {
          throw new Error(
            `Export job for ${args.reportName} was accepted but returned no job id`,
          );
        }

        const deadline = Date.now() + args.timeoutSeconds * 1000;
        let current = job;
        let status = str(current.status);
        while (
          status !== "completed" && status !== "failed" && Date.now() < deadline
        ) {
          await new Promise((r) => setTimeout(r, 5000));
          const { data: polled } = await graphRequest(
            context.globalArgs,
            "GET",
            `${base}/deviceManagement/reports/exportJobs('${
              encodeURIComponent(jobId)
            }')`,
          );
          current = (polled ?? {}) as Record<string, unknown>;
          status = str(current.status);
        }

        if (status === "failed") {
          throw new Error(
            `Export job ${jobId} for ${args.reportName} failed server-side`,
          );
        }
        if (status !== "completed") {
          throw new Error(
            `Export job ${jobId} for ${args.reportName} still ${
              status || "pending"
            } after ${args.timeoutSeconds}s — raise timeoutSeconds or narrow the report`,
          );
        }

        context.logger.info("Export job {id} completed", { id: jobId });
        // Named by report only, not by job id: each report keeps one current
        // export resource that later runs refresh in place, rather than
        // accumulating a new instance per run behind a truncated job id.
        const handle = await context.writeResource(
          "reportExport",
          slugify(args.reportName, "export"),
          current,
        );
        return { dataHandles: [handle] };
      },
    },
  },
};
