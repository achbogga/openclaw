export type ProviderId = "google" | "microsoft";

export type ProviderRuntimeConfig = {
  id: ProviderId;
  accessToken: string;
  calendarEnabled: boolean;
  mailEnabled: boolean;
  calendarIds: string[];
  writableCalendarIds: string[];
  userId?: string;
};

export type ExecutiveAssistantRuntimeConfig = {
  timezone?: string;
  lookaheadDays: number;
  maxCalendarResults: number;
  maxMailResults: number;
  providers: ProviderRuntimeConfig[];
};

export type CalendarEvent = {
  provider: ProviderId;
  calendarId: string;
  id: string;
  title: string;
  startTime?: string;
  endTime?: string;
  isAllDay: boolean;
  location?: string;
  attendees: string[];
  url?: string;
  writable: boolean;
};

export type MailSearchResult = {
  provider: ProviderId;
  id: string;
  threadId: string;
  subject: string;
  from?: string;
  receivedAt?: string;
  snippet?: string;
  isRead?: boolean;
  url?: string;
};

export type MailThreadMessage = {
  id: string;
  subject: string;
  from?: string;
  to: string[];
  cc: string[];
  receivedAt?: string;
  isRead?: boolean;
  bodyText?: string;
  snippet?: string;
  url?: string;
};

export type MailThread = {
  provider: ProviderId;
  threadId: string;
  subject: string;
  messages: MailThreadMessage[];
};
