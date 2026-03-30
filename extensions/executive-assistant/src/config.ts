import type { OpenClawConfig } from "openclaw/plugin-sdk/config-runtime";
import {
  normalizeResolvedSecretInputString,
  normalizeSecretInput,
} from "openclaw/plugin-sdk/secret-input";
import type {
  ExecutiveAssistantRuntimeConfig,
  ProviderId,
  ProviderRuntimeConfig,
} from "./types.js";

export const DEFAULT_LOOKAHEAD_DAYS = 3;
export const DEFAULT_MAX_CALENDAR_RESULTS = 25;
export const DEFAULT_MAX_MAIL_RESULTS = 10;
export const GOOGLE_ACCESS_TOKEN_ENV = "OPENCLAW_EXECUTIVE_ASSISTANT_GOOGLE_ACCESS_TOKEN";
export const MICROSOFT_ACCESS_TOKEN_ENV = "OPENCLAW_EXECUTIVE_ASSISTANT_MICROSOFT_ACCESS_TOKEN";

type RawProviderConfig = {
  enabled?: unknown;
  accessToken?: unknown;
  calendarEnabled?: unknown;
  mailEnabled?: unknown;
  calendarIds?: unknown;
  writableCalendarIds?: unknown;
  userId?: unknown;
};

type RawDefaultsConfig = {
  timezone?: unknown;
  lookaheadDays?: unknown;
  maxCalendarResults?: unknown;
  maxMailResults?: unknown;
};

type RawPluginConfig = {
  defaults?: unknown;
  google?: unknown;
  microsoft?: unknown;
};

function asRecord(value: unknown): Record<string, unknown> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  return value as Record<string, unknown>;
}

function normalizeStringArray(value: unknown, fallback: string[]): string[] {
  if (!Array.isArray(value)) {
    return fallback;
  }
  const normalized = value
    .filter((entry) => typeof entry === "string")
    .map((entry) => entry.trim())
    .filter(Boolean);
  return normalized.length > 0 ? normalized : fallback;
}

function normalizePositiveInt(value: unknown, fallback: number, min: number, max: number): number {
  const raw =
    typeof value === "number"
      ? value
      : typeof value === "string" && value.trim()
        ? Number.parseInt(value, 10)
        : Number.NaN;
  if (!Number.isFinite(raw)) {
    return fallback;
  }
  return Math.max(min, Math.min(max, Math.trunc(raw)));
}

function readConfiguredToken(value: unknown, path: string, envVarName: string): string | undefined {
  const configured = normalizeSecretInput(
    normalizeResolvedSecretInputString({
      value,
      path,
    }),
  );
  if (configured) {
    return configured;
  }
  return normalizeSecretInput(process.env[envVarName]) || undefined;
}

function resolveProviderConfig(params: {
  id: ProviderId;
  raw: RawProviderConfig | undefined;
  accessTokenEnv: string;
  defaultCalendarIds: string[];
}): ProviderRuntimeConfig | null {
  const raw = params.raw;
  const enabled = raw?.enabled === false ? false : true;
  const accessToken = readConfiguredToken(
    raw?.accessToken,
    `plugins.entries.executive-assistant.config.${params.id}.accessToken`,
    params.accessTokenEnv,
  );

  if (!enabled || !accessToken) {
    return null;
  }

  const calendarEnabled = raw?.calendarEnabled === false ? false : true;
  const mailEnabled = raw?.mailEnabled === false ? false : true;

  return {
    id: params.id,
    accessToken,
    calendarEnabled,
    mailEnabled,
    calendarIds: normalizeStringArray(raw?.calendarIds, params.defaultCalendarIds),
    writableCalendarIds: normalizeStringArray(raw?.writableCalendarIds, []),
    userId:
      typeof raw?.userId === "string" && raw.userId.trim().length > 0
        ? raw.userId.trim()
        : undefined,
  };
}

export function resolveExecutiveAssistantRuntimeConfig(
  cfg?: OpenClawConfig,
): ExecutiveAssistantRuntimeConfig {
  const pluginConfig = asRecord(cfg?.plugins?.entries?.["executive-assistant"]?.config) as
    | RawPluginConfig
    | undefined;
  const defaults = asRecord(pluginConfig?.defaults) as RawDefaultsConfig | undefined;

  const google = resolveProviderConfig({
    id: "google",
    raw: asRecord(pluginConfig?.google) as RawProviderConfig | undefined,
    accessTokenEnv: GOOGLE_ACCESS_TOKEN_ENV,
    defaultCalendarIds: ["primary"],
  });
  const microsoft = resolveProviderConfig({
    id: "microsoft",
    raw: asRecord(pluginConfig?.microsoft) as RawProviderConfig | undefined,
    accessTokenEnv: MICROSOFT_ACCESS_TOKEN_ENV,
    defaultCalendarIds: ["default"],
  });

  const timezone =
    (typeof defaults?.timezone === "string" && defaults.timezone.trim()) ||
    (typeof cfg?.agents?.defaults?.userTimezone === "string" &&
    cfg.agents.defaults.userTimezone.trim()
      ? cfg.agents.defaults.userTimezone.trim()
      : undefined);

  return {
    timezone: timezone || undefined,
    lookaheadDays: normalizePositiveInt(defaults?.lookaheadDays, DEFAULT_LOOKAHEAD_DAYS, 1, 30),
    maxCalendarResults: normalizePositiveInt(
      defaults?.maxCalendarResults,
      DEFAULT_MAX_CALENDAR_RESULTS,
      1,
      100,
    ),
    maxMailResults: normalizePositiveInt(
      defaults?.maxMailResults,
      DEFAULT_MAX_MAIL_RESULTS,
      1,
      100,
    ),
    providers: [google, microsoft].filter((provider): provider is ProviderRuntimeConfig =>
      Boolean(provider),
    ),
  };
}

export function requireProviderConfig(
  config: ExecutiveAssistantRuntimeConfig,
  providerId: ProviderId,
): ProviderRuntimeConfig {
  const provider = config.providers.find((entry) => entry.id === providerId);
  if (!provider) {
    throw new Error(
      `Provider "${providerId}" is not configured. Add an access token under plugins.entries.executive-assistant.config.${providerId}.accessToken.`,
    );
  }
  return provider;
}
