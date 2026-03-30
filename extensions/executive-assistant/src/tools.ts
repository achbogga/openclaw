import { Type } from "@sinclair/typebox";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk/plugin-runtime";
import {
  jsonResult,
  readNumberParam,
  readStringArrayParam,
  readStringParam,
} from "openclaw/plugin-sdk/provider-web-search";
import { resolveExecutiveAssistantRuntimeConfig, requireProviderConfig } from "./config.js";
import {
  createGoogleCalendarEvent,
  getGoogleMailThread,
  listGoogleCalendarEvents,
  listGoogleUnreadMail,
  searchGoogleMail,
} from "./google.js";
import {
  createMicrosoftCalendarEvent,
  getMicrosoftMailThread,
  listMicrosoftCalendarEvents,
  listMicrosoftUnreadMail,
  searchMicrosoftMail,
} from "./microsoft.js";
import type {
  CalendarEvent,
  ExecutiveAssistantRuntimeConfig,
  MailSearchResult,
  ProviderId,
  ProviderRuntimeConfig,
} from "./types.js";

const DAY_MS = 24 * 60 * 60 * 1000;
const PROVIDER_IDS = ["google", "microsoft"] as const;

function optionalStringEnum<const T extends readonly string[]>(values: T, description: string) {
  return Type.Optional(
    Type.Unsafe<T[number]>({
      type: "string",
      enum: [...values],
      description,
    }),
  );
}

const CalendarWindowSchema = Type.Object(
  {
    provider: optionalStringEnum(
      PROVIDER_IDS,
      "Optional provider filter. Omit to query every configured provider.",
    ),
    start_time: Type.Optional(
      Type.String({
        description: "ISO date/time window start. Defaults to now for event listing.",
      }),
    ),
    end_time: Type.Optional(
      Type.String({
        description: "ISO date/time window end. Defaults to start + configured lookahead window.",
      }),
    ),
    days: Type.Optional(
      Type.Number({
        description: "Fallback window length in days when end_time is omitted.",
        minimum: 1,
        maximum: 30,
      }),
    ),
    max_results: Type.Optional(
      Type.Number({
        description: "Maximum number of events to return per provider.",
        minimum: 1,
        maximum: 100,
      }),
    ),
    calendar_ids: Type.Optional(
      Type.Array(Type.String(), {
        description:
          "Optional calendar ids for a single provider. When using multiple providers, rely on configured calendarIds instead.",
      }),
    ),
  },
  { additionalProperties: false },
);

const ConflictSchema = Type.Object(
  {
    provider: optionalStringEnum(
      PROVIDER_IDS,
      "Optional provider filter. Omit to check every configured provider.",
    ),
    start_time: Type.String({
      description: "Proposed event start as ISO date/time.",
    }),
    end_time: Type.String({
      description: "Proposed event end as ISO date/time.",
    }),
    calendar_ids: Type.Optional(
      Type.Array(Type.String(), {
        description:
          "Optional calendar ids for a single provider. When using multiple providers, rely on configured calendarIds instead.",
      }),
    ),
  },
  { additionalProperties: false },
);

const CreateEventSchema = Type.Object(
  {
    provider: optionalStringEnum(
      PROVIDER_IDS,
      "Provider to write to. Required when more than one provider has writable calendars.",
    ),
    calendar_id: Type.Optional(
      Type.String({
        description:
          "Writable calendar id to create the event on. Required when the chosen provider exposes more than one writable calendar.",
      }),
    ),
    title: Type.String({
      description: "Event title.",
    }),
    start_time: Type.String({
      description: "Event start as ISO date/time.",
    }),
    end_time: Type.String({
      description: "Event end as ISO date/time.",
    }),
    description: Type.Optional(
      Type.String({
        description: "Optional plain-text event description.",
      }),
    ),
    attendees: Type.Optional(
      Type.Array(Type.String(), {
        description: "Optional attendee email addresses.",
      }),
    ),
    confirm: Type.Optional(
      Type.Boolean({
        description: "Must be true only after the user explicitly confirms the calendar write.",
      }),
    ),
  },
  { additionalProperties: false },
);

const MailSearchSchema = Type.Object(
  {
    provider: optionalStringEnum(
      PROVIDER_IDS,
      "Optional provider filter. Omit to query every configured mail provider.",
    ),
    query: Type.String({
      description:
        "Mail query. Gmail receives the raw Gmail query string; Microsoft Graph uses keyword-style filtering with support for is:unread, from:, subject:, after:, and before: tokens.",
    }),
    max_results: Type.Optional(
      Type.Number({
        description: "Maximum number of messages to return per provider.",
        minimum: 1,
        maximum: 100,
      }),
    ),
  },
  { additionalProperties: false },
);

const MailThreadSchema = Type.Object(
  {
    provider: optionalStringEnum(
      PROVIDER_IDS,
      "Provider for the thread. Omit only when exactly one mail provider is configured.",
    ),
    thread_id: Type.String({
      description: "Provider-native thread id (Gmail thread id or Microsoft conversationId).",
    }),
  },
  { additionalProperties: false },
);

const BriefingSchema = Type.Object(
  {
    date: Type.Optional(
      Type.String({
        description:
          "Optional local calendar date in YYYY-MM-DD form. Defaults to the current local date.",
      }),
    ),
    include_mail: Type.Optional(
      Type.Boolean({
        description: "Whether to include unread mail in the briefing. Defaults to true.",
      }),
    ),
    max_mail_results: Type.Optional(
      Type.Number({
        description: "Maximum unread mail items to include per provider.",
        minimum: 1,
        maximum: 25,
      }),
    ),
  },
  { additionalProperties: false },
);

function parseIso(value: string, label: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    throw new Error(`${label} must be a valid ISO date or date-time`);
  }
  return date.toISOString();
}

function sortEvents(events: CalendarEvent[]): CalendarEvent[] {
  return [...events].sort((left, right) => {
    const leftTime = left.startTime ? new Date(left.startTime).getTime() : 0;
    const rightTime = right.startTime ? new Date(right.startTime).getTime() : 0;
    return leftTime - rightTime;
  });
}

function sortMail(messages: MailSearchResult[]): MailSearchResult[] {
  return [...messages].sort((left, right) => {
    const leftTime = left.receivedAt ? new Date(left.receivedAt).getTime() : 0;
    const rightTime = right.receivedAt ? new Date(right.receivedAt).getTime() : 0;
    return rightTime - leftTime;
  });
}

function requireConfiguredProviders(
  config: ExecutiveAssistantRuntimeConfig,
  capability: "calendar" | "mail",
  providerId?: string,
): ProviderRuntimeConfig[] {
  const providers =
    providerId && providerId.trim()
      ? [requireProviderConfig(config, providerId as ProviderId)]
      : config.providers;
  const filtered = providers.filter((provider) =>
    capability === "calendar" ? provider.calendarEnabled : provider.mailEnabled,
  );
  if (filtered.length === 0) {
    throw new Error(
      capability === "calendar"
        ? "No calendar providers are configured for executive-assistant."
        : "No mail providers are configured for executive-assistant.",
    );
  }
  return filtered;
}

function resolveCalendarWindow(
  rawParams: Record<string, unknown>,
  config: ExecutiveAssistantRuntimeConfig,
  defaultStart: "now" | "day-start" = "now",
): { startTime: string; endTime: string } {
  const startParam = readStringParam(rawParams, "start_time");
  const endParam = readStringParam(rawParams, "end_time");
  const days = readNumberParam(rawParams, "days", { integer: true }) ?? config.lookaheadDays;

  const startTime =
    startParam ??
    (defaultStart === "day-start"
      ? new Date(new Date().toDateString()).toISOString()
      : new Date().toISOString());
  const normalizedStart = parseIso(startTime, "start_time");
  const normalizedEnd = endParam
    ? parseIso(endParam, "end_time")
    : new Date(new Date(normalizedStart).getTime() + days * DAY_MS).toISOString();
  if (new Date(normalizedEnd).getTime() <= new Date(normalizedStart).getTime()) {
    throw new Error("end_time must be later than start_time");
  }
  return { startTime: normalizedStart, endTime: normalizedEnd };
}

function resolveDailyWindow(rawDate?: string): {
  date: string;
  startTime: string;
  endTime: string;
} {
  if (rawDate) {
    const trimmed = rawDate.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      throw new Error("date must use YYYY-MM-DD format");
    }
    const start = new Date(`${trimmed}T00:00:00`);
    return {
      date: trimmed,
      startTime: start.toISOString(),
      endTime: new Date(start.getTime() + DAY_MS).toISOString(),
    };
  }

  const start = new Date(new Date().toDateString());
  return {
    date: start.toISOString().slice(0, 10),
    startTime: start.toISOString(),
    endTime: new Date(start.getTime() + DAY_MS).toISOString(),
  };
}

async function listCalendarEvents(params: {
  providers: ProviderRuntimeConfig[];
  startTime: string;
  endTime: string;
  maxResults: number;
  calendarIds?: string[];
}): Promise<CalendarEvent[]> {
  const results = await Promise.all(
    params.providers.map((provider) => {
      if (provider.id === "google") {
        return listGoogleCalendarEvents({
          provider,
          startTime: params.startTime,
          endTime: params.endTime,
          maxResults: params.maxResults,
          calendarIds: params.calendarIds,
        });
      }
      return listMicrosoftCalendarEvents({
        provider,
        startTime: params.startTime,
        endTime: params.endTime,
        maxResults: params.maxResults,
        calendarIds: params.calendarIds,
      });
    }),
  );
  return sortEvents(results.flat());
}

function findOverlaps(events: CalendarEvent[]): Array<{
  left: CalendarEvent;
  right: CalendarEvent;
}> {
  const sorted = sortEvents(events);
  const overlaps: Array<{ left: CalendarEvent; right: CalendarEvent }> = [];
  for (let index = 0; index < sorted.length - 1; index += 1) {
    const left = sorted[index];
    const right = sorted[index + 1];
    if (!left?.endTime || !right?.startTime) {
      continue;
    }
    if (new Date(left.endTime).getTime() > new Date(right.startTime).getTime()) {
      overlaps.push({ left, right });
    }
  }
  return overlaps;
}

function resolveWritableTarget(params: {
  config: ExecutiveAssistantRuntimeConfig;
  providerId?: string;
  calendarId?: string;
}): { provider: ProviderRuntimeConfig; calendarId: string } {
  const providers = params.providerId?.trim()
    ? [requireProviderConfig(params.config, params.providerId as ProviderId)]
    : params.config.providers;
  const writableProviders = providers.filter(
    (provider) => provider.calendarEnabled && provider.writableCalendarIds.length > 0,
  );

  if (writableProviders.length === 0) {
    throw new Error("No writable calendars are configured for executive-assistant.");
  }

  if (params.providerId?.trim()) {
    const provider = writableProviders[0];
    if (!provider) {
      throw new Error(`Provider "${params.providerId}" does not expose writable calendars.`);
    }
    const calendarId = params.calendarId?.trim();
    if (calendarId) {
      if (!provider.writableCalendarIds.includes(calendarId)) {
        throw new Error(
          `Calendar "${calendarId}" is not allowlisted for writes on ${provider.id}.`,
        );
      }
      return { provider, calendarId };
    }
    if (provider.writableCalendarIds.length !== 1) {
      throw new Error(
        `Provider "${provider.id}" has multiple writable calendars. Specify calendar_id explicitly.`,
      );
    }
    return { provider, calendarId: provider.writableCalendarIds[0] ?? "" };
  }

  if (params.calendarId?.trim()) {
    const matchingProviders = writableProviders.filter((provider) =>
      provider.writableCalendarIds.includes(params.calendarId!.trim()),
    );
    if (matchingProviders.length === 0) {
      throw new Error(`Calendar "${params.calendarId}" is not allowlisted for writes.`);
    }
    if (matchingProviders.length > 1) {
      throw new Error(
        `Calendar "${params.calendarId}" exists on multiple writable providers. Specify provider explicitly.`,
      );
    }
    return { provider: matchingProviders[0]!, calendarId: params.calendarId.trim() };
  }

  if (writableProviders.length !== 1) {
    throw new Error("Multiple writable providers are configured. Specify provider explicitly.");
  }

  const provider = writableProviders[0]!;
  if (provider.writableCalendarIds.length !== 1) {
    throw new Error(
      `Provider "${provider.id}" has multiple writable calendars. Specify calendar_id explicitly.`,
    );
  }
  return { provider, calendarId: provider.writableCalendarIds[0] ?? "" };
}

async function searchMail(params: {
  providers: ProviderRuntimeConfig[];
  query: string;
  maxResults: number;
}): Promise<MailSearchResult[]> {
  const results = await Promise.all(
    params.providers.map((provider) => {
      if (provider.id === "google") {
        return searchGoogleMail({
          provider,
          query: params.query,
          maxResults: params.maxResults,
        });
      }
      return searchMicrosoftMail({
        provider,
        query: params.query,
        maxResults: params.maxResults,
      });
    }),
  );
  return sortMail(results.flat());
}

async function listUnreadMail(params: {
  providers: ProviderRuntimeConfig[];
  maxResults: number;
}): Promise<MailSearchResult[]> {
  const results = await Promise.all(
    params.providers.map((provider) => {
      if (provider.id === "google") {
        return listGoogleUnreadMail({ provider, maxResults: params.maxResults });
      }
      return listMicrosoftUnreadMail({ provider, maxResults: params.maxResults });
    }),
  );
  return sortMail(results.flat());
}

export function createExecutiveAssistantTools(api: OpenClawPluginApi) {
  return [
    {
      name: "calendar_list_events",
      label: "Calendar List Events",
      description:
        "Read Google Calendar and Microsoft calendar events across the configured accounts. Use this before proposing times or creating an event.",
      parameters: CalendarWindowSchema,
      execute: async (_toolCallId: string, rawParams: Record<string, unknown>) => {
        const config = resolveExecutiveAssistantRuntimeConfig(api.config);
        const providerId = readStringParam(rawParams, "provider");
        const calendarIds = readStringArrayParam(rawParams, "calendar_ids");
        if (calendarIds && !providerId) {
          throw new Error(
            "calendar_ids requires provider so the target calendars are unambiguous.",
          );
        }
        const providers = requireConfiguredProviders(config, "calendar", providerId);
        const maxResults =
          readNumberParam(rawParams, "max_results", { integer: true }) ?? config.maxCalendarResults;
        const window = resolveCalendarWindow(rawParams, config);
        const events = await listCalendarEvents({
          providers,
          startTime: window.startTime,
          endTime: window.endTime,
          maxResults,
          calendarIds,
        });
        return jsonResult({
          startTime: window.startTime,
          endTime: window.endTime,
          events,
        });
      },
    },
    {
      name: "calendar_find_conflicts",
      label: "Calendar Find Conflicts",
      description:
        "Check whether a proposed time collides with any configured Google or Microsoft calendar event.",
      parameters: ConflictSchema,
      execute: async (_toolCallId: string, rawParams: Record<string, unknown>) => {
        const config = resolveExecutiveAssistantRuntimeConfig(api.config);
        const providerId = readStringParam(rawParams, "provider");
        const calendarIds = readStringArrayParam(rawParams, "calendar_ids");
        if (calendarIds && !providerId) {
          throw new Error(
            "calendar_ids requires provider so the target calendars are unambiguous.",
          );
        }
        const startTime = parseIso(
          readStringParam(rawParams, "start_time", { required: true }),
          "start_time",
        );
        const endTime = parseIso(
          readStringParam(rawParams, "end_time", { required: true }),
          "end_time",
        );
        if (new Date(endTime).getTime() <= new Date(startTime).getTime()) {
          throw new Error("end_time must be later than start_time");
        }
        const providers = requireConfiguredProviders(config, "calendar", providerId);
        const events = await listCalendarEvents({
          providers,
          startTime,
          endTime,
          maxResults: config.maxCalendarResults,
          calendarIds,
        });
        return jsonResult({
          ok: events.length === 0,
          proposed: { startTime, endTime },
          conflicts: events,
        });
      },
    },
    {
      name: "calendar_create_personal_event",
      label: "Calendar Create Personal Event",
      description:
        "Create an event only on explicitly allowlisted personal calendars. Ask the user for confirmation first, then call this tool with confirm=true.",
      parameters: CreateEventSchema,
      execute: async (_toolCallId: string, rawParams: Record<string, unknown>) => {
        const confirm = rawParams.confirm === true;
        if (!confirm) {
          throw new Error(
            "calendar_create_personal_event requires confirm=true after the user explicitly approves the write.",
          );
        }
        const config = resolveExecutiveAssistantRuntimeConfig(api.config);
        const providerId = readStringParam(rawParams, "provider");
        const calendarId = readStringParam(rawParams, "calendar_id");
        const target = resolveWritableTarget({ config, providerId, calendarId });
        const title = readStringParam(rawParams, "title", { required: true });
        const startTime = parseIso(
          readStringParam(rawParams, "start_time", { required: true }),
          "start_time",
        );
        const endTime = parseIso(
          readStringParam(rawParams, "end_time", { required: true }),
          "end_time",
        );
        if (new Date(endTime).getTime() <= new Date(startTime).getTime()) {
          throw new Error("end_time must be later than start_time");
        }
        const description = readStringParam(rawParams, "description");
        const attendees = readStringArrayParam(rawParams, "attendees");

        const event =
          target.provider.id === "google"
            ? await createGoogleCalendarEvent({
                provider: target.provider,
                calendarId: target.calendarId,
                title,
                startTime,
                endTime,
                description,
                attendees,
              })
            : await createMicrosoftCalendarEvent({
                provider: target.provider,
                calendarId: target.calendarId,
                title,
                startTime,
                endTime,
                description,
                attendees,
              });

        return jsonResult({
          confirmed: true,
          event,
        });
      },
    },
    {
      name: "mail_search_readonly",
      label: "Mail Search Readonly",
      description:
        "Search Gmail and Microsoft mail in read-only mode. Use it to find recent threads before reading a full thread.",
      parameters: MailSearchSchema,
      execute: async (_toolCallId: string, rawParams: Record<string, unknown>) => {
        const config = resolveExecutiveAssistantRuntimeConfig(api.config);
        const query = readStringParam(rawParams, "query", { required: true });
        const providerId = readStringParam(rawParams, "provider");
        const providers = requireConfiguredProviders(config, "mail", providerId);
        const maxResults =
          readNumberParam(rawParams, "max_results", { integer: true }) ?? config.maxMailResults;
        const messages = await searchMail({
          providers,
          query,
          maxResults,
        });
        return jsonResult({
          query,
          messages,
        });
      },
    },
    {
      name: "mail_get_thread",
      label: "Mail Get Thread",
      description:
        "Fetch a full read-only mail thread by provider-native thread id (Gmail thread id or Microsoft conversationId).",
      parameters: MailThreadSchema,
      execute: async (_toolCallId: string, rawParams: Record<string, unknown>) => {
        const config = resolveExecutiveAssistantRuntimeConfig(api.config);
        const threadId = readStringParam(rawParams, "thread_id", { required: true });
        const providerId = readStringParam(rawParams, "provider");
        const providers = requireConfiguredProviders(config, "mail", providerId);
        if (providers.length !== 1) {
          throw new Error(
            "mail_get_thread requires provider when more than one mail provider is configured.",
          );
        }
        const provider = providers[0]!;
        const thread =
          provider.id === "google"
            ? await getGoogleMailThread({ provider, threadId })
            : await getMicrosoftMailThread({ provider, threadId });
        return jsonResult(thread);
      },
    },
    {
      name: "briefing_daily",
      label: "Briefing Daily",
      description:
        "Build a daily briefing using today's calendar agenda plus unread mail across configured Google and Microsoft accounts.",
      parameters: BriefingSchema,
      execute: async (_toolCallId: string, rawParams: Record<string, unknown>) => {
        const config = resolveExecutiveAssistantRuntimeConfig(api.config);
        const { date, startTime, endTime } = resolveDailyWindow(readStringParam(rawParams, "date"));
        const includeMail = rawParams.include_mail !== false;
        const calendarProviders = requireConfiguredProviders(config, "calendar");
        const events = await listCalendarEvents({
          providers: calendarProviders,
          startTime,
          endTime,
          maxResults: config.maxCalendarResults,
        });
        const overlaps = findOverlaps(events).map((pair) => ({
          left: {
            id: pair.left.id,
            title: pair.left.title,
            provider: pair.left.provider,
            startTime: pair.left.startTime,
            endTime: pair.left.endTime,
          },
          right: {
            id: pair.right.id,
            title: pair.right.title,
            provider: pair.right.provider,
            startTime: pair.right.startTime,
            endTime: pair.right.endTime,
          },
        }));

        const unreadMail =
          includeMail && config.providers.some((provider) => provider.mailEnabled)
            ? await listUnreadMail({
                providers: requireConfiguredProviders(config, "mail"),
                maxResults:
                  readNumberParam(rawParams, "max_mail_results", { integer: true }) ??
                  Math.min(config.maxMailResults, 5),
              })
            : [];

        return jsonResult({
          date,
          timezone: config.timezone,
          startTime,
          endTime,
          events,
          unreadMail,
          conflicts: overlaps,
        });
      },
    },
  ];
}
