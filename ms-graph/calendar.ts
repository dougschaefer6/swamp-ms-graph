import { z } from "npm:zod@4.3.6";
import {
  type CalendarEvent,
  type DataHandle,
  GRAPH_BASE,
  graphList,
  graphRequest,
  mapEvent,
  type MethodContext,
  MsGraphGlobalArgsSchema,
  slugify,
} from "./_client.ts";

/**
 * `@dougschaefer/ms-graph-calendar` model — wraps Microsoft Graph v1.0 calendar
 * endpoints to read room and user mailbox schedules.
 *
 * This is the calendar member of the broad `@dougschaefer/ms-graph-*` family. It
 * powers the Identity-Aware Room Services (IARS) meeting-agent: the agent calls
 * `getCurrentMeeting` or `getNextMeeting` on a room mailbox to learn what meeting
 * is in progress (or coming up), who is on the invite, and what the subject is —
 * so the room can prepare ahead of time (load the right AV scene, surface context,
 * greet attendees by name).
 *
 * Authentication uses the shared app-only client-credentials client from
 * `_client.ts` against the configured Entra app vault. The required application
 * permission is `Calendars.Read` (admin-consented and live-verified in the tenant):
 *
 *   client_id:     ${{ vault.get(azure-asei, client_id) }}
 *   client_secret: ${{ vault.get(azure-asei, client_secret) }}
 *   tenant_id:     ${{ vault.get(azure-asei, tenant_id) }}
 */
export const model = {
  type: "@dougschaefer/ms-graph-calendar",
  version: "2026.07.13.1",
  globalArguments: MsGraphGlobalArgsSchema,
  resources: {
    calendarView: {
      description:
        "Ordered list of meetings for a mailbox within a time window (from /users/{mailbox}/calendarView). Normalized to {subject, start, end, organizer, attendees, attendeeCount}.",
      schema: z.object({
        mailbox: z.string(),
        start: z.string(),
        end: z.string(),
        meetings: z.array(z.unknown()),
        count: z.number(),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "5m",
      garbageCollection: 20,
    },
    event: {
      description:
        "A single calendar event fetched by id from /users/{mailbox}/events/{id}, normalized to the standard meeting shape.",
      schema: z.object({
        mailbox: z.string(),
        id: z.string(),
        meeting: z.union([z.unknown(), z.null()]),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "5m",
      garbageCollection: 20,
    },
    currentMeeting: {
      description:
        "The meeting in progress right now for the mailbox, or null if none. Derived from a narrow calendarView window centered on the current time. The IARS meeting-agent's primary calendar signal.",
      schema: z.object({
        mailbox: z.string(),
        found: z.boolean(),
        meeting: z.union([z.unknown(), z.null()]),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "2m",
      garbageCollection: 30,
    },
    nextMeeting: {
      description:
        "The next upcoming meeting for the mailbox (the first meeting starting after now). Null if nothing is on the calendar in the look-ahead window. Used by the IARS meeting-agent to prepare the room ahead of time.",
      schema: z.object({
        mailbox: z.string(),
        found: z.boolean(),
        meeting: z.union([z.unknown(), z.null()]),
        minutesUntilStart: z.number(),
        capturedAt: z.iso.datetime(),
      }),
      lifetime: "2m",
      garbageCollection: 30,
    },
  },
  methods: {
    listEvents: {
      description:
        "GET /users/{roomEmail}/calendarView — return all meetings for a room or user mailbox within a caller-supplied time window, ordered by start time. " +
        "Normalized output: {subject, start, end, organizer, attendees:[{name,email,type}], attendeeCount}. " +
        "The Prefer: outlook.timezone header ensures start/end are returned in the mailbox's local timezone (default: Eastern Standard Time). " +
        "Use this for bulk schedule inspection; use getCurrentMeeting / getNextMeeting for the IARS meeting-agent's hot path. (Permission: Calendars.Read)",
      arguments: z.object({
        roomEmail: z.string().describe(
          "Room or user mailbox address, e.g. conf-b@example.com",
        ),
        startDateTime: z.string().describe(
          "Window start in ISO-8601 format, e.g. 2026-06-29T00:00:00",
        ),
        endDateTime: z.string().describe(
          "Window end in ISO-8601 format, e.g. 2026-06-29T23:59:59",
        ),
        top: z.number().int().default(50).describe(
          "Max events per page (1-999). The method auto-follows @odata.nextLink.",
        ),
      }),
      execute: async (
        args: {
          roomEmail: string;
          startDateTime: string;
          endDateTime: string;
          top: number;
        },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info(
          "calendarView for {room} [{start} -> {end}]",
          {
            room: args.roomEmail,
            start: args.startDateTime,
            end: args.endDateTime,
          },
        );
        const select =
          "subject,start,end,organizer,attendees,isAllDay,isCancelled";
        const url =
          `${GRAPH_BASE}/users/${
            encodeURIComponent(args.roomEmail)
          }/calendarView` +
          `?startDateTime=${encodeURIComponent(args.startDateTime)}` +
          `&endDateTime=${encodeURIComponent(args.endDateTime)}` +
          `&$select=${select}` +
          `&$orderby=start/dateTime` +
          `&$top=${Math.min(999, Math.max(1, args.top))}`;
        const prefer = `outlook.timezone="${context.globalArgs.timezone}"`;
        const raw = await graphList(context.globalArgs, url, { prefer });
        const meetings: CalendarEvent[] = [];
        for (const ev of raw) {
          const mapped = ev as Record<string, unknown>;
          if (!mapped.isCancelled) meetings.push(mapEvent(ev));
        }
        const instanceName = `${slugify(args.roomEmail)}-${
          args.startDateTime.slice(0, 10)
        }`;
        const handle = await context.writeResource(
          "calendarView",
          instanceName,
          {
            mailbox: args.roomEmail,
            start: args.startDateTime,
            end: args.endDateTime,
            meetings,
            count: meetings.length,
            capturedAt: new Date().toISOString(),
          },
        );
        context.logger.info("Found {n} meeting(s) for {room}", {
          n: meetings.length,
          room: args.roomEmail,
        });
        return { dataHandles: [handle] };
      },
    },

    getEvent: {
      description:
        "GET /users/{roomEmail}/events/{id} — fetch a single calendar event by its Graph event id, normalized to the standard meeting shape. (Permission: Calendars.Read)",
      arguments: z.object({
        roomEmail: z.string().describe(
          "Room or user mailbox address that owns the event, e.g. conf-b@example.com",
        ),
        eventId: z.string().describe("Graph event id"),
      }),
      execute: async (
        args: { roomEmail: string; eventId: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        context.logger.info("getEvent {id} for {room}", {
          id: args.eventId,
          room: args.roomEmail,
        });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${encodeURIComponent(args.roomEmail)}/events/${
            encodeURIComponent(args.eventId)
          }`,
          { prefer: `outlook.timezone="${context.globalArgs.timezone}"` },
        );
        const handle = await context.writeResource(
          "event",
          `${slugify(args.roomEmail)}-${slugify(args.eventId)}`,
          {
            mailbox: args.roomEmail,
            id: args.eventId,
            meeting: data ? mapEvent(data) : null,
            capturedAt: new Date().toISOString(),
          },
        );
        return { dataHandles: [handle] };
      },
    },

    getCurrentMeeting: {
      description:
        "Convenience method: returns the meeting currently in progress for the room, or null if the room is free. " +
        "Queries calendarView with a narrow window around now and returns the first non-cancelled event whose start <= now < end. " +
        "This is the IARS meeting-agent's primary calendar signal: it tells the room what meeting is happening right now, who the attendees are, and what the subject is. (Permission: Calendars.Read)",
      arguments: z.object({
        roomEmail: z.string().describe(
          "Room mailbox address, e.g. conf-b@example.com",
        ),
      }),
      execute: async (
        args: { roomEmail: string },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const now = new Date();
        // Widen the window slightly so a meeting that just started (< 1 min ago)
        // is caught even if the clock skews a bit.
        const windowStart = new Date(now.getTime() - 2 * 60 * 1000);
        const windowEnd = new Date(now.getTime() + 1 * 60 * 1000);
        const fmt = (d: Date): string =>
          d.toISOString().replace(/\.\d{3}Z$/, "");
        context.logger.info("getCurrentMeeting for {room}", {
          room: args.roomEmail,
        });
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${
            encodeURIComponent(args.roomEmail)
          }/calendarView` +
            `?startDateTime=${encodeURIComponent(fmt(windowStart))}` +
            `&endDateTime=${encodeURIComponent(fmt(windowEnd))}` +
            `&$select=subject,start,end,organizer,attendees,isAllDay,isCancelled` +
            `&$orderby=start/dateTime` +
            `&$top=10`,
          { prefer: `outlook.timezone="${context.globalArgs.timezone}"` },
        );
        const d = (data ?? {}) as Record<string, unknown>;
        const events = Array.isArray(d.value) ? d.value : [];
        const nowMs = now.getTime();
        let current: CalendarEvent | null = null;
        for (const ev of events) {
          const evObj = (ev ?? {}) as Record<string, unknown>;
          if (evObj.isCancelled) continue;
          const mapped = mapEvent(ev);
          const startMs = mapped.start ? new Date(mapped.start).getTime() : 0;
          const endMs = mapped.end ? new Date(mapped.end).getTime() : 0;
          if (startMs <= nowMs && nowMs < endMs) {
            current = mapped;
            break;
          }
        }
        const handle = await context.writeResource(
          "currentMeeting",
          slugify(args.roomEmail),
          {
            mailbox: args.roomEmail,
            found: current !== null,
            meeting: current,
            capturedAt: now.toISOString(),
          },
        );
        if (current) {
          context.logger.info("Current meeting: {subject} ({n} attendee(s))", {
            subject: current.subject,
            n: current.attendeeCount,
          });
        } else {
          context.logger.info("Room is free (no current meeting)", {});
        }
        return { dataHandles: [handle] };
      },
    },

    getNextMeeting: {
      description:
        "Convenience method: returns the next upcoming meeting for the room — the first meeting that starts after now (within the look-ahead window). " +
        "Returns null if nothing is on the calendar. Also returns minutesUntilStart so the IARS meeting-agent can decide how early to prepare the room. " +
        "Pair with getCurrentMeeting: if the room is free, use getNextMeeting to know what to stage. (Permission: Calendars.Read)",
      arguments: z.object({
        roomEmail: z.string().describe(
          "Room mailbox address, e.g. conf-b@example.com",
        ),
        lookAheadHours: z.number().default(8).describe(
          "How many hours ahead to search for the next meeting (default 8 — remainder of the workday)",
        ),
      }),
      execute: async (
        args: { roomEmail: string; lookAheadHours: number },
        context: MethodContext,
      ): Promise<{ dataHandles: DataHandle[] }> => {
        const now = new Date();
        const windowEnd = new Date(
          now.getTime() + args.lookAheadHours * 3600 * 1000,
        );
        const fmt = (d: Date): string =>
          d.toISOString().replace(/\.\d{3}Z$/, "");
        context.logger.info(
          "getNextMeeting for {room} (look-ahead {h}h)",
          { room: args.roomEmail, h: args.lookAheadHours },
        );
        const { data } = await graphRequest(
          context.globalArgs,
          "GET",
          `${GRAPH_BASE}/users/${
            encodeURIComponent(args.roomEmail)
          }/calendarView` +
            `?startDateTime=${encodeURIComponent(fmt(now))}` +
            `&endDateTime=${encodeURIComponent(fmt(windowEnd))}` +
            `&$select=subject,start,end,organizer,attendees,isAllDay,isCancelled` +
            `&$orderby=start/dateTime` +
            `&$top=10`,
          { prefer: `outlook.timezone="${context.globalArgs.timezone}"` },
        );
        const d = (data ?? {}) as Record<string, unknown>;
        const events = Array.isArray(d.value) ? d.value : [];
        const nowMs = now.getTime();
        let next: CalendarEvent | null = null;
        let minutesUntilStart = 0;
        for (const ev of events) {
          const evObj = (ev ?? {}) as Record<string, unknown>;
          if (evObj.isCancelled) continue;
          const mapped = mapEvent(ev);
          const startMs = mapped.start ? new Date(mapped.start).getTime() : 0;
          if (startMs > nowMs) {
            next = mapped;
            minutesUntilStart = Math.round((startMs - nowMs) / 60000);
            break;
          }
        }
        const handle = await context.writeResource(
          "nextMeeting",
          slugify(args.roomEmail),
          {
            mailbox: args.roomEmail,
            found: next !== null,
            meeting: next,
            minutesUntilStart,
            capturedAt: now.toISOString(),
          },
        );
        if (next) {
          context.logger.info(
            "Next meeting: {subject} in {m} minute(s)",
            { subject: next.subject, m: minutesUntilStart },
          );
        } else {
          context.logger.info("No upcoming meetings in the next {h}h", {
            h: args.lookAheadHours,
          });
        }
        return { dataHandles: [handle] };
      },
    },
  },
};
