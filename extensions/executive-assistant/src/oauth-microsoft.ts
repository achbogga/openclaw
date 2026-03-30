import type { OAuthCredential } from "openclaw/plugin-sdk/agent-runtime";
import type { ProviderAuthContext, ProviderAuthResult } from "openclaw/plugin-sdk/provider-auth";
import { ensureGlobalUndiciEnvProxyDispatcher, sleep } from "openclaw/plugin-sdk/runtime-env";
import { buildAuthProfileId } from "../../../src/agents/auth-profiles/identity.js";
import {
  buildExecutiveAssistantConfigPatch,
  MICROSOFT_AUTH_PROVIDER_ID,
  MICROSOFT_OAUTH_CLIENT_ID_ENV,
  MICROSOFT_OAUTH_CLIENT_SECRET_ENV,
  MICROSOFT_OAUTH_TENANT_ID_ENV,
} from "./config.js";
import type { ProviderId } from "./types.js";

const MICROSOFT_GRAPH_ME_URL =
  "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName";
const MICROSOFT_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Calendars.ReadWrite",
];

type MicrosoftDeviceCodeResponse = {
  device_code?: string;
  user_code?: string;
  verification_uri?: string;
  expires_in?: number;
  interval?: number;
  message?: string;
  error?: string;
  error_description?: string;
};

type MicrosoftTokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

type MicrosoftIdentity = {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
};

type StoredMicrosoftOAuthCredential = OAuthCredential & {
  clientId?: string;
  clientSecret?: string;
  tenantId?: string;
};

function trimOrUndefined(value: string | undefined | null): string | undefined {
  const trimmed = value?.trim();
  return trimmed ? trimmed : undefined;
}

function truncateErrorText(value: string): string {
  const normalized = value.replace(/\s+/g, " ").trim();
  if (!normalized) {
    return "";
  }
  return normalized.length > 280 ? `${normalized.slice(0, 277)}...` : normalized;
}

function deviceCodeUrl(tenantId: string): string {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/devicecode`;
}

function tokenUrl(tenantId: string): string {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
}

async function fetchMicrosoftIdentity(accessToken: string): Promise<MicrosoftIdentity> {
  ensureGlobalUndiciEnvProxyDispatcher();

  const response = await fetch(MICROSOFT_GRAPH_ME_URL, {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
    },
  });
  const raw = await response.text();
  if (!response.ok) {
    throw new Error(
      `Microsoft Graph identity failed: ${truncateErrorText(raw) || response.status}`,
    );
  }
  return JSON.parse(raw) as MicrosoftIdentity;
}

async function requestMicrosoftDeviceCode(params: {
  clientId: string;
  tenantId: string;
}): Promise<
  Required<
    Pick<
      MicrosoftDeviceCodeResponse,
      "device_code" | "user_code" | "verification_uri" | "expires_in"
    >
  > &
    Pick<MicrosoftDeviceCodeResponse, "interval" | "message">
> {
  ensureGlobalUndiciEnvProxyDispatcher();

  const body = new URLSearchParams({
    client_id: params.clientId,
    scope: MICROSOFT_SCOPES.join(" "),
  });
  const response = await fetch(deviceCodeUrl(params.tenantId), {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const raw = await response.text();
  if (!response.ok) {
    throw new Error(
      `Microsoft device-code request failed: ${truncateErrorText(raw) || response.status}`,
    );
  }
  const payload = JSON.parse(raw) as MicrosoftDeviceCodeResponse;
  if (
    !payload.device_code ||
    !payload.user_code ||
    !payload.verification_uri ||
    !payload.expires_in
  ) {
    throw new Error(
      payload.error_description?.trim() ||
        "Microsoft device-code response did not include the required fields.",
    );
  }
  return {
    device_code: payload.device_code,
    user_code: payload.user_code,
    verification_uri: payload.verification_uri,
    expires_in: payload.expires_in,
    ...(payload.interval ? { interval: payload.interval } : {}),
    ...(payload.message ? { message: payload.message } : {}),
  };
}

async function pollMicrosoftAccessToken(params: {
  clientId: string;
  clientSecret?: string;
  tenantId: string;
  deviceCode: string;
  intervalMs: number;
  expiresAt: number;
}): Promise<{ access: string; refresh: string; expires: number }> {
  ensureGlobalUndiciEnvProxyDispatcher();

  while (Date.now() < params.expiresAt) {
    const body = new URLSearchParams({
      client_id: params.clientId,
      device_code: params.deviceCode,
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      ...(params.clientSecret ? { client_secret: params.clientSecret } : {}),
    });
    const response = await fetch(tokenUrl(params.tenantId), {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    });

    const raw = await response.text();
    if (!response.ok && response.status >= 500) {
      throw new Error(
        `Microsoft token polling failed: ${truncateErrorText(raw) || response.status}`,
      );
    }

    const payload = JSON.parse(raw) as MicrosoftTokenResponse;
    if (payload.access_token && payload.refresh_token && payload.expires_in) {
      return {
        access: payload.access_token,
        refresh: payload.refresh_token,
        expires: Date.now() + payload.expires_in * 1000 - 60_000,
      };
    }

    const error = payload.error?.trim();
    if (error === "authorization_pending") {
      await sleep(params.intervalMs);
      continue;
    }
    if (error === "slow_down") {
      await sleep(params.intervalMs + 5_000);
      continue;
    }
    if (error === "authorization_declined" || error === "access_denied") {
      throw new Error("Microsoft login was cancelled.");
    }
    if (error === "expired_token" || error === "bad_verification_code") {
      throw new Error("Microsoft device code expired. Run login again.");
    }
    throw new Error(
      payload.error_description?.trim() ||
        `Microsoft device-code flow failed${error ? `: ${error}` : "."}`,
    );
  }

  throw new Error("Microsoft device code expired. Run login again.");
}

async function refreshMicrosoftOAuthCredential(
  credential: StoredMicrosoftOAuthCredential,
): Promise<StoredMicrosoftOAuthCredential> {
  const clientId =
    trimOrUndefined(credential.clientId) ??
    trimOrUndefined(process.env[MICROSOFT_OAUTH_CLIENT_ID_ENV]);
  const tenantId =
    trimOrUndefined(credential.tenantId) ??
    trimOrUndefined(process.env[MICROSOFT_OAUTH_TENANT_ID_ENV]) ??
    "common";
  const clientSecret =
    trimOrUndefined(credential.clientSecret) ??
    trimOrUndefined(process.env[MICROSOFT_OAUTH_CLIENT_SECRET_ENV]);

  if (!clientId) {
    throw new Error(
      `Microsoft OAuth refresh requires clientId on the auth profile or ${MICROSOFT_OAUTH_CLIENT_ID_ENV}.`,
    );
  }
  if (!credential.refresh?.trim()) {
    throw new Error("Microsoft OAuth refresh token is missing.");
  }

  ensureGlobalUndiciEnvProxyDispatcher();

  const body = new URLSearchParams({
    client_id: clientId,
    refresh_token: credential.refresh,
    grant_type: "refresh_token",
    scope: MICROSOFT_SCOPES.join(" "),
    ...(clientSecret ? { client_secret: clientSecret } : {}),
  });
  const response = await fetch(tokenUrl(tenantId), {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const raw = await response.text();
  if (!response.ok) {
    throw new Error(`Microsoft token refresh failed: ${truncateErrorText(raw) || response.status}`);
  }
  const payload = JSON.parse(raw) as MicrosoftTokenResponse;
  if (!payload.access_token || !payload.expires_in) {
    throw new Error(
      payload.error_description?.trim() ||
        "Microsoft token refresh did not return access_token and expires_in.",
    );
  }

  return {
    ...credential,
    type: "oauth",
    provider: MICROSOFT_AUTH_PROVIDER_ID,
    access: payload.access_token,
    refresh: payload.refresh_token?.trim() || credential.refresh,
    expires: Date.now() + payload.expires_in * 1000 - 60_000,
    clientId,
    tenantId,
    ...(clientSecret ? { clientSecret } : {}),
  };
}

async function promptMicrosoftClientConfig(
  ctx: ProviderAuthContext,
): Promise<{ clientId: string; clientSecret?: string; tenantId: string }> {
  const envClientId = trimOrUndefined(process.env[MICROSOFT_OAUTH_CLIENT_ID_ENV]);
  const envClientSecret = trimOrUndefined(process.env[MICROSOFT_OAUTH_CLIENT_SECRET_ENV]);
  const envTenantId = trimOrUndefined(process.env[MICROSOFT_OAUTH_TENANT_ID_ENV]);

  const clientId =
    envClientId ??
    String(
      await ctx.prompter.text({
        message: "Enter Microsoft Entra app client id",
        placeholder: "00000000-0000-0000-0000-000000000000",
        validate: (value) => (value.trim() ? undefined : "Required"),
      }),
    ).trim();
  const tenantId = String(
    await ctx.prompter.text({
      message: "Enter Microsoft tenant id",
      initialValue: envTenantId ?? "common",
      placeholder: "common",
      validate: (value) => (value.trim() ? undefined : "Required"),
    }),
  ).trim();
  const clientSecret =
    envClientSecret ??
    trimOrUndefined(
      String(
        await ctx.prompter.text({
          message: "Enter Microsoft client secret (optional)",
          placeholder: "Leave blank for public/native clients",
        }),
      ),
    );

  return {
    clientId,
    tenantId,
    ...(clientSecret ? { clientSecret } : {}),
  };
}

function buildAuthResult(params: {
  logicalProviderId: ProviderId;
  access: string;
  refresh: string;
  expires: number;
  email?: string;
  displayName?: string;
  accountId?: string;
  clientId: string;
  clientSecret?: string;
  tenantId: string;
}): ProviderAuthResult {
  const profileId = buildAuthProfileId({
    providerId: MICROSOFT_AUTH_PROVIDER_ID,
    profileName: params.email ?? params.displayName ?? "default",
  });
  const credential: OAuthCredential = {
    type: "oauth",
    provider: MICROSOFT_AUTH_PROVIDER_ID,
    access: params.access,
    refresh: params.refresh,
    expires: params.expires,
    ...(params.email ? { email: params.email } : {}),
    ...(params.displayName ? { displayName: params.displayName } : {}),
    ...(params.accountId ? { accountId: params.accountId } : {}),
    clientId: params.clientId,
    tenantId: params.tenantId,
    ...(params.clientSecret ? { clientSecret: params.clientSecret } : {}),
  } as OAuthCredential;

  return {
    profiles: [{ profileId, credential }],
    configPatch: buildExecutiveAssistantConfigPatch({
      providerId: params.logicalProviderId,
      authProfileId: profileId,
      userId: "me",
    }),
    notes: [
      `Microsoft account: ${params.email ?? params.displayName ?? "connected"}`,
      "Mail stays read-only. Calendar writes still require writableCalendarIds plus confirm=true.",
      `Tenant: ${params.tenantId}`,
    ],
  };
}

export async function runExecutiveAssistantMicrosoftOAuth(
  ctx: ProviderAuthContext,
): Promise<ProviderAuthResult> {
  const { clientId, clientSecret, tenantId } = await promptMicrosoftClientConfig(ctx);

  await ctx.prompter.note(
    [
      "This flow uses Microsoft device-code login for Graph Calendar + Mail.",
      "Your Entra app should allow device-code/public-client auth and Graph delegated scopes for User.Read, Mail.Read, Calendars.ReadWrite, and offline_access.",
      "",
      `Tenant: ${tenantId}`,
    ].join("\n"),
    "Executive Assistant Microsoft OAuth",
  );

  const progress = ctx.prompter.progress("Requesting Microsoft device code…");
  let deviceCodeProgressActive = true;
  try {
    const device = await requestMicrosoftDeviceCode({ clientId, tenantId });
    progress.stop("Microsoft device code ready");
    deviceCodeProgressActive = false;

    if (device.message?.trim()) {
      ctx.runtime.log(`\n${device.message.trim()}\n`);
    } else {
      ctx.runtime.log(
        `\nVisit ${device.verification_uri} and enter code ${device.user_code} to continue.\n`,
      );
    }
    if (!ctx.isRemote) {
      try {
        await ctx.openUrl(device.verification_uri);
      } catch {
        // The verification URL is already logged for manual use.
      }
    }

    const polling = ctx.prompter.progress("Waiting for Microsoft authorization…");
    const tokens = await pollMicrosoftAccessToken({
      clientId,
      ...(clientSecret ? { clientSecret } : {}),
      tenantId,
      deviceCode: device.device_code,
      intervalMs: Math.max(5_000, (device.interval ?? 5) * 1_000),
      expiresAt: Date.now() + device.expires_in * 1_000,
    });
    polling.update("Reading Microsoft account identity…");
    const identity = await fetchMicrosoftIdentity(tokens.access);
    polling.stop("Microsoft OAuth complete");

    return buildAuthResult({
      logicalProviderId: "microsoft",
      ...tokens,
      email: trimOrUndefined(identity.mail) ?? trimOrUndefined(identity.userPrincipalName),
      displayName: trimOrUndefined(identity.displayName),
      accountId: trimOrUndefined(identity.id),
      clientId,
      tenantId,
      ...(clientSecret ? { clientSecret } : {}),
    });
  } catch (error) {
    if (deviceCodeProgressActive) {
      progress.stop("Microsoft OAuth failed");
    }
    throw error;
  }
}

export function buildExecutiveAssistantMicrosoftProvider() {
  return {
    id: MICROSOFT_AUTH_PROVIDER_ID,
    label: "Executive Assistant Microsoft",
    envVars: [
      MICROSOFT_OAUTH_CLIENT_ID_ENV,
      MICROSOFT_OAUTH_TENANT_ID_ENV,
      MICROSOFT_OAUTH_CLIENT_SECRET_ENV,
    ],
    auth: [
      {
        id: "device-code",
        label: "Microsoft device code",
        hint: "Connect Microsoft Graph Calendar + Mail from the terminal",
        kind: "device_code" as const,
        wizard: {
          choiceId: "executive-assistant-microsoft",
          choiceLabel: "Executive Assistant Microsoft",
          choiceHint: "Connect Microsoft Calendar + Mail",
          groupId: "executive-assistant",
          groupLabel: "Executive Assistant",
          groupHint: "Calendar + mail connectors",
        },
        run: async (ctx: ProviderAuthContext) => await runExecutiveAssistantMicrosoftOAuth(ctx),
      },
    ],
    refreshOAuth: async (credential: OAuthCredential) =>
      await refreshMicrosoftOAuthCredential(credential as StoredMicrosoftOAuthCredential),
  };
}
