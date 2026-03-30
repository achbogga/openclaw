import { fetchJson } from "./http.js";
import type {
  CalendarEvent,
  MailSearchResult,
  MailThread,
  MailThreadMessage,
  ResolvedProviderRuntimeConfig,
} from "./types.js";

type GoogleCalendarEvent = {
  id?: string;
  summary?: string;
  htmlLink?: string;
  location?: string;
  start?: {
    dateTime?: string;
    date?: string;
  };
  end?: {
    dateTime?: string;
    date?: string;
  };
  attendees?: Array<{ email?: string }>;
};

type GoogleCalendarEventsResponse = {
  items?: GoogleCalendarEvent[];
};

type GoogleMessageListResponse = {
  messages?: Array<{ id?: string; threadId?: string }>;
};

type GoogleMessage = {
  id?: string;
  threadId?: string;
  internalDate?: string;
  snippet?: string;
  labelIds?: string[];
  payload?: {
    mimeType?: string;
    headers?: Array<{ name?: string; value?: string }>;
    body?: { data?: string };
    parts?: Array<GoogleMessage["payload"]>;
  };
};

type GoogleThreadResponse = {
  id?: string;
  messages?: GoogleMessage[];
};

const GOOGLE_CALENDAR_ROOT = "https://www.googleapis.com/calendar/v3";
const GOOGLE_GMAIL_ROOT = "https://gmail.googleapis.com/gmail/v1/users/me";

function encodeCalendarId(calendarId: string): string {
  return encodeURIComponent(calendarId);
}

function headerValue(message: GoogleMessage, name: string): string | undefined {
  const headers = message.payload?.headers;
  if (!Array.isArray(headers)) {
    return undefined;
  }
  const lower = name.toLowerCase();
  const match = headers.find((entry) => entry.name?.toLowerCase() === lower);
  const value = match?.value?.trim();
  return value || undefined;
}

function decodeBase64Url(value: string): string {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  return Buffer.from(padded, "base64").toString("utf8");
}

function stripHtml(value: string): string {
  return value
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractGoogleBodyText(payload: GoogleMessage["payload"] | undefined): string | undefined {
  if (!payload) {
    return undefined;
  }

  if (payload.mimeType === "text/plain" && typeof payload.body?.data === "string") {
    return decodeBase64Url(payload.body.data).trim() || undefined;
  }

  if (Array.isArray(payload.parts)) {
    for (const part of payload.parts) {
      const text = extractGoogleBodyText(part);
      if (text) {
        return text;
      }
    }
  }

  if (payload.mimeType === "text/html" && typeof payload.body?.data === "string") {
    const html = decodeBase64Url(payload.body.data);
    const text = stripHtml(html);
    return text || undefined;
  }

  if (typeof payload.body?.data === "string") {
    const text = decodeBase64Url(payload.body.data).trim();
    return text || undefined;
  }

  return undefined;
}

function normalizeGoogleEvent(
  provider: ResolvedProviderRuntimeConfig,
  calendarId: string,
  event: GoogleCalendarEvent,
): CalendarEvent {
  return {
    provider: "google",
    calendarId,
    id: event.id ?? "",
    title: event.summary?.trim() || "(untitled)",
    startTime: event.start?.dateTime ?? event.start?.date,
    endTime: event.end?.dateTime ?? event.end?.date,
    isAllDay: Boolean(event.start?.date && !event.start?.dateTime),
    location: event.location?.trim() || undefined,
    attendees:
      event.attendees
        ?.map((attendee) => attendee.email?.trim())
        .filter((value): value is string => Boolean(value)) ?? [],
    url: event.htmlLink?.trim() || undefined,
    writable: provider.writableCalendarIds.includes(calendarId),
  };
}

function normalizeGoogleMessageSummary(message: GoogleMessage): MailSearchResult {
  return {
    provider: "google",
    id: message.id ?? "",
    threadId: message.threadId ?? "",
    subject: headerValue(message, "Subject") || "(no subject)",
    from: headerValue(message, "From"),
    receivedAt: message.internalDate
      ? new Date(Number.parseInt(message.internalDate, 10)).toISOString()
      : headerValue(message, "Date"),
    snippet: message.snippet?.trim() || undefined,
    isRead: !(message.labelIds ?? []).includes("UNREAD"),
    url: message.id ? `https://mail.google.com/mail/u/0/#all/${message.id}` : undefined,
  };
}

function normalizeGoogleThreadMessage(message: GoogleMessage): MailThreadMessage {
  const to = headerValue(message, "To")
    ?.split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  const cc = headerValue(message, "Cc")
    ?.split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  return {
    id: message.id ?? "",
    subject: headerValue(message, "Subject") || "(no subject)",
    from: headerValue(message, "From"),
    to: to ?? [],
    cc: cc ?? [],
    receivedAt: message.internalDate
      ? new Date(Number.parseInt(message.internalDate, 10)).toISOString()
      : headerValue(message, "Date"),
    isRead: !(message.labelIds ?? []).includes("UNREAD"),
    bodyText: extractGoogleBodyText(message.payload),
    snippet: message.snippet?.trim() || undefined,
    url: message.id ? `https://mail.google.com/mail/u/0/#all/${message.id}` : undefined,
  };
}

export async function listGoogleCalendarEvents(params: {
  provider: ResolvedProviderRuntimeConfig;
  startTime: string;
  endTime: string;
  maxResults: number;
  calendarIds?: string[];
}): Promise<CalendarEvent[]> {
  const calendarIds = params.calendarIds?.length ? params.calendarIds : params.provider.calendarIds;
  const results = await Promise.all(
    calendarIds.map(async (calendarId) => {
      const query = new URLSearchParams({
        singleEvents: "true",
        orderBy: "startTime",
        timeMin: params.startTime,
        timeMax: params.endTime,
        maxResults: String(params.maxResults),
      });
      const payload = await fetchJson<GoogleCalendarEventsResponse>({
        url: `${GOOGLE_CALENDAR_ROOT}/calendars/${encodeCalendarId(calendarId)}/events?${query.toString()}`,
        token: params.provider.accessToken,
        label: `Google Calendar ${calendarId}`,
      });
      return (payload.items ?? []).map((event) =>
        normalizeGoogleEvent(params.provider, calendarId, event),
      );
    }),
  );
  return results.flat();
}

export async function createGoogleCalendarEvent(params: {
  provider: ResolvedProviderRuntimeConfig;
  calendarId: string;
  title: string;
  startTime: string;
  endTime: string;
  description?: string;
  attendees?: string[];
}): Promise<CalendarEvent> {
  const created = await fetchJson<GoogleCalendarEvent>({
    url: `${GOOGLE_CALENDAR_ROOT}/calendars/${encodeCalendarId(params.calendarId)}/events`,
    token: params.provider.accessToken,
    label: `Google Calendar create ${params.calendarId}`,
    method: "POST",
    body: {
      summary: params.title,
      ...(params.description ? { description: params.description } : {}),
      start: { dateTime: params.startTime },
      end: { dateTime: params.endTime },
      ...(params.attendees?.length
        ? { attendees: params.attendees.map((email) => ({ email })) }
        : {}),
    },
  });
  return normalizeGoogleEvent(params.provider, params.calendarId, created);
}

export async function searchGoogleMail(params: {
  provider: ResolvedProviderRuntimeConfig;
  query: string;
  maxResults: number;
}): Promise<MailSearchResult[]> {
  const query = new URLSearchParams({
    q: params.query,
    maxResults: String(params.maxResults),
  });
  const listing = await fetchJson<GoogleMessageListResponse>({
    url: `${GOOGLE_GMAIL_ROOT}/messages?${query.toString()}`,
    token: params.provider.accessToken,
    label: "Gmail search",
  });
  const ids = (listing.messages ?? [])
    .map((entry) => entry.id?.trim())
    .filter((value): value is string => Boolean(value));
  const messages = await Promise.all(
    ids.map((id) =>
      fetchJson<GoogleMessage>({
        url: `${GOOGLE_GMAIL_ROOT}/messages/${encodeURIComponent(id)}?format=metadata&metadataHeaders=From&metadataHeaders=Subject&metadataHeaders=Date`,
        token: params.provider.accessToken,
        label: `Gmail message ${id}`,
      }),
    ),
  );
  return messages.map(normalizeGoogleMessageSummary);
}

export async function listGoogleUnreadMail(params: {
  provider: ResolvedProviderRuntimeConfig;
  maxResults: number;
}): Promise<MailSearchResult[]> {
  return await searchGoogleMail({
    provider: params.provider,
    query: "is:unread newer_than:7d",
    maxResults: params.maxResults,
  });
}

export async function getGoogleMailThread(params: {
  provider: ResolvedProviderRuntimeConfig;
  threadId: string;
}): Promise<MailThread> {
  const thread = await fetchJson<GoogleThreadResponse>({
    url: `${GOOGLE_GMAIL_ROOT}/threads/${encodeURIComponent(params.threadId)}?format=full`,
    token: params.provider.accessToken,
    label: `Gmail thread ${params.threadId}`,
  });
  const messages = (thread.messages ?? []).map(normalizeGoogleThreadMessage);
  return {
    provider: "google",
    threadId: thread.id ?? params.threadId,
    subject: messages[0]?.subject ?? "(no subject)",
    messages,
  };
}
