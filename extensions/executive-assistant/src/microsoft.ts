import { fetchJson } from "./http.js";
import type {
  CalendarEvent,
  MailSearchResult,
  MailThread,
  MailThreadMessage,
  ResolvedProviderRuntimeConfig,
} from "./types.js";

type GraphDateTime = {
  dateTime?: string;
  timeZone?: string;
};

type GraphCalendarEvent = {
  id?: string;
  subject?: string;
  webLink?: string;
  location?: {
    displayName?: string;
  };
  start?: GraphDateTime;
  end?: GraphDateTime;
  attendees?: Array<{
    emailAddress?: {
      address?: string;
    };
  }>;
  isAllDay?: boolean;
};

type GraphCalendarEventsResponse = {
  value?: GraphCalendarEvent[];
};

type GraphMessage = {
  id?: string;
  conversationId?: string;
  subject?: string;
  webLink?: string;
  receivedDateTime?: string;
  isRead?: boolean;
  bodyPreview?: string;
  from?: {
    emailAddress?: {
      address?: string;
      name?: string;
    };
  };
  toRecipients?: Array<{
    emailAddress?: {
      address?: string;
    };
  }>;
  ccRecipients?: Array<{
    emailAddress?: {
      address?: string;
    };
  }>;
  body?: {
    content?: string;
  };
};

type GraphMessagesResponse = {
  value?: GraphMessage[];
};

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const GRAPH_PREFER_HEADERS = {
  Prefer: 'outlook.timezone="UTC", outlook.body-content-type="text"',
};

type ParsedMailQuery = {
  unreadOnly: boolean;
  from: string[];
  subject: string[];
  before?: string;
  after?: string;
  freeText: string[];
};

function graphUserPath(provider: ResolvedProviderRuntimeConfig): string {
  if (provider.userId && provider.userId !== "me") {
    return `/users/${encodeURIComponent(provider.userId)}`;
  }
  return "/me";
}

function normalizeGraphDateTime(value: GraphDateTime | undefined): string | undefined {
  const raw = value?.dateTime?.trim();
  if (!raw) {
    return undefined;
  }
  if (/[zZ]$|[+-]\d{2}:\d{2}$/.test(raw)) {
    return new Date(raw).toISOString();
  }
  if ((value?.timeZone ?? "").toUpperCase() === "UTC") {
    return new Date(`${raw}Z`).toISOString();
  }
  return raw;
}

function toGraphUtcDateTime(input: string): string {
  return new Date(input).toISOString().replace(/\.\d{3}Z$/, "");
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function formatGraphEmail(message: GraphMessage["from"]): string | undefined {
  const email = message?.emailAddress?.address?.trim();
  const name = message?.emailAddress?.name?.trim();
  if (!email) {
    return undefined;
  }
  if (name && name.toLowerCase() !== email.toLowerCase()) {
    return `${name} <${email}>`;
  }
  return email;
}

function normalizeGraphEvent(
  provider: ResolvedProviderRuntimeConfig,
  calendarId: string,
  event: GraphCalendarEvent,
): CalendarEvent {
  return {
    provider: "microsoft",
    calendarId,
    id: event.id ?? "",
    title: event.subject?.trim() || "(untitled)",
    startTime: normalizeGraphDateTime(event.start),
    endTime: normalizeGraphDateTime(event.end),
    isAllDay: event.isAllDay === true,
    location: event.location?.displayName?.trim() || undefined,
    attendees:
      event.attendees
        ?.map((attendee) => attendee.emailAddress?.address?.trim())
        .filter((value): value is string => Boolean(value)) ?? [],
    url: event.webLink?.trim() || undefined,
    writable: provider.writableCalendarIds.includes(calendarId),
  };
}

function normalizeGraphMessageSummary(message: GraphMessage): MailSearchResult {
  return {
    provider: "microsoft",
    id: message.id ?? "",
    threadId: message.conversationId ?? "",
    subject: message.subject?.trim() || "(no subject)",
    from: formatGraphEmail(message.from),
    receivedAt: message.receivedDateTime?.trim() || undefined,
    snippet: message.bodyPreview?.trim() || undefined,
    isRead: message.isRead,
    url: message.webLink?.trim() || undefined,
  };
}

function normalizeGraphThreadMessage(message: GraphMessage): MailThreadMessage {
  return {
    id: message.id ?? "",
    subject: message.subject?.trim() || "(no subject)",
    from: formatGraphEmail(message.from),
    to:
      message.toRecipients
        ?.map((entry) => entry.emailAddress?.address?.trim())
        .filter((value): value is string => Boolean(value)) ?? [],
    cc:
      message.ccRecipients
        ?.map((entry) => entry.emailAddress?.address?.trim())
        .filter((value): value is string => Boolean(value)) ?? [],
    receivedAt: message.receivedDateTime?.trim() || undefined,
    isRead: message.isRead,
    bodyText: message.body?.content?.trim() || undefined,
    snippet: message.bodyPreview?.trim() || undefined,
    url: message.webLink?.trim() || undefined,
  };
}

function parseMailQuery(query: string): ParsedMailQuery {
  const parsed: ParsedMailQuery = {
    unreadOnly: false,
    from: [],
    subject: [],
    freeText: [],
  };

  for (const token of query
    .split(/\s+/)
    .map((value) => value.trim())
    .filter(Boolean)) {
    const lower = token.toLowerCase();
    if (lower === "is:unread") {
      parsed.unreadOnly = true;
      continue;
    }
    if (lower.startsWith("from:")) {
      const value = token.slice(5).trim().toLowerCase();
      if (value) {
        parsed.from.push(value);
      }
      continue;
    }
    if (lower.startsWith("subject:")) {
      const value = token.slice(8).trim().toLowerCase();
      if (value) {
        parsed.subject.push(value);
      }
      continue;
    }
    if (lower.startsWith("before:")) {
      parsed.before = token.slice(7).trim();
      continue;
    }
    if (lower.startsWith("after:")) {
      parsed.after = token.slice(6).trim();
      continue;
    }
    parsed.freeText.push(lower);
  }

  return parsed;
}

function matchesGraphMessage(message: MailSearchResult, parsed: ParsedMailQuery): boolean {
  const haystack =
    `${message.subject} ${message.from ?? ""} ${message.snippet ?? ""}`.toLowerCase();
  if (parsed.unreadOnly && message.isRead === true) {
    return false;
  }
  if (
    parsed.from.length > 0 &&
    !parsed.from.every((value) => (message.from ?? "").toLowerCase().includes(value))
  ) {
    return false;
  }
  if (
    parsed.subject.length > 0 &&
    !parsed.subject.every((value) => message.subject.toLowerCase().includes(value))
  ) {
    return false;
  }
  if (parsed.after && message.receivedAt && new Date(message.receivedAt) < new Date(parsed.after)) {
    return false;
  }
  if (
    parsed.before &&
    message.receivedAt &&
    new Date(message.receivedAt) > new Date(parsed.before)
  ) {
    return false;
  }
  return parsed.freeText.every((value) => haystack.includes(value));
}

export async function listMicrosoftCalendarEvents(params: {
  provider: ResolvedProviderRuntimeConfig;
  startTime: string;
  endTime: string;
  maxResults: number;
  calendarIds?: string[];
}): Promise<CalendarEvent[]> {
  const calendarIds = params.calendarIds?.length ? params.calendarIds : params.provider.calendarIds;
  const userPath = graphUserPath(params.provider);
  const results = await Promise.all(
    calendarIds.map(async (calendarId) => {
      const path =
        calendarId === "default" || calendarId === "primary"
          ? `${userPath}/calendar/calendarView`
          : `${userPath}/calendars/${encodeURIComponent(calendarId)}/calendarView`;
      const query = new URLSearchParams({
        startDateTime: params.startTime,
        endDateTime: params.endTime,
        $top: String(params.maxResults),
        $select: "id,subject,webLink,location,start,end,attendees,isAllDay",
      });
      const payload = await fetchJson<GraphCalendarEventsResponse>({
        url: `${GRAPH_ROOT}${path}?${query.toString()}`,
        token: params.provider.accessToken,
        label: `Microsoft calendar ${calendarId}`,
        headers: GRAPH_PREFER_HEADERS,
      });
      return (payload.value ?? []).map((event) =>
        normalizeGraphEvent(params.provider, calendarId, event),
      );
    }),
  );
  return results.flat();
}

export async function createMicrosoftCalendarEvent(params: {
  provider: ResolvedProviderRuntimeConfig;
  calendarId: string;
  title: string;
  startTime: string;
  endTime: string;
  description?: string;
  attendees?: string[];
}): Promise<CalendarEvent> {
  const userPath = graphUserPath(params.provider);
  const path =
    params.calendarId === "default" || params.calendarId === "primary"
      ? `${userPath}/calendar/events`
      : `${userPath}/calendars/${encodeURIComponent(params.calendarId)}/events`;
  const created = await fetchJson<GraphCalendarEvent>({
    url: `${GRAPH_ROOT}${path}`,
    token: params.provider.accessToken,
    label: `Microsoft calendar create ${params.calendarId}`,
    method: "POST",
    headers: GRAPH_PREFER_HEADERS,
    body: {
      subject: params.title,
      ...(params.description ? { body: { contentType: "text", content: params.description } } : {}),
      start: { dateTime: toGraphUtcDateTime(params.startTime), timeZone: "UTC" },
      end: { dateTime: toGraphUtcDateTime(params.endTime), timeZone: "UTC" },
      ...(params.attendees?.length
        ? {
            attendees: params.attendees.map((email) => ({
              emailAddress: { address: email },
              type: "required",
            })),
          }
        : {}),
    },
  });
  return normalizeGraphEvent(params.provider, params.calendarId, created);
}

export async function searchMicrosoftMail(params: {
  provider: ResolvedProviderRuntimeConfig;
  query: string;
  maxResults: number;
}): Promise<MailSearchResult[]> {
  const userPath = graphUserPath(params.provider);
  const parsed = parseMailQuery(params.query);
  const query = new URLSearchParams({
    $top: String(Math.max(params.maxResults * 3, 15)),
    $select: "id,conversationId,subject,webLink,receivedDateTime,isRead,bodyPreview,from",
    $orderby: "receivedDateTime DESC",
  });
  if (parsed.unreadOnly) {
    query.set("$filter", "isRead eq false");
  }
  const payload = await fetchJson<GraphMessagesResponse>({
    url: `${GRAPH_ROOT}${userPath}/messages?${query.toString()}`,
    token: params.provider.accessToken,
    label: "Microsoft mail search",
    headers: GRAPH_PREFER_HEADERS,
  });
  return (payload.value ?? [])
    .map(normalizeGraphMessageSummary)
    .filter((message) => matchesGraphMessage(message, parsed))
    .slice(0, params.maxResults);
}

export async function listMicrosoftUnreadMail(params: {
  provider: ResolvedProviderRuntimeConfig;
  maxResults: number;
}): Promise<MailSearchResult[]> {
  return await searchMicrosoftMail({
    provider: params.provider,
    query: "is:unread",
    maxResults: params.maxResults,
  });
}

export async function getMicrosoftMailThread(params: {
  provider: ResolvedProviderRuntimeConfig;
  threadId: string;
}): Promise<MailThread> {
  const userPath = graphUserPath(params.provider);
  const query = new URLSearchParams({
    $filter: `conversationId eq '${escapeODataString(params.threadId)}'`,
    $select:
      "id,conversationId,subject,webLink,receivedDateTime,isRead,bodyPreview,body,from,toRecipients,ccRecipients",
    $orderby: "receivedDateTime asc",
  });
  const payload = await fetchJson<GraphMessagesResponse>({
    url: `${GRAPH_ROOT}${userPath}/messages?${query.toString()}`,
    token: params.provider.accessToken,
    label: `Microsoft thread ${params.threadId}`,
    headers: GRAPH_PREFER_HEADERS,
  });
  const messages = (payload.value ?? []).map(normalizeGraphThreadMessage);
  return {
    provider: "microsoft",
    threadId: params.threadId,
    subject: messages[0]?.subject ?? "(no subject)",
    messages,
  };
}
