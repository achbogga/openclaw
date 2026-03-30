import { randomBytes } from "node:crypto";
import { createServer } from "node:http";
import type { OAuthCredential } from "openclaw/plugin-sdk/agent-runtime";
import type { ProviderAuthContext, ProviderAuthResult } from "openclaw/plugin-sdk/provider-auth";
import { generatePkceVerifierChallenge, toFormUrlEncoded } from "openclaw/plugin-sdk/provider-auth";
import { ensureGlobalUndiciEnvProxyDispatcher, isWSL2Sync } from "openclaw/plugin-sdk/runtime-env";
import { buildAuthProfileId } from "../../../src/agents/auth-profiles/identity.js";
import {
  buildExecutiveAssistantConfigPatch,
  GOOGLE_ACCESS_TOKEN_ENV,
  GOOGLE_AUTH_PROVIDER_ID,
  GOOGLE_OAUTH_CLIENT_ID_ENV,
  GOOGLE_OAUTH_CLIENT_SECRET_ENV,
} from "./config.js";
import type { ProviderId } from "./types.js";

const GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const GOOGLE_USERINFO_URL = "https://www.googleapis.com/oauth2/v1/userinfo?alt=json";
const GOOGLE_REDIRECT_URI = "http://localhost:8085/oauth2callback";
const GOOGLE_SCOPES = [
  "openid",
  "email",
  "profile",
  "https://www.googleapis.com/auth/calendar",
  "https://www.googleapis.com/auth/gmail.readonly",
];

type GoogleTokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

type GoogleUserInfo = {
  id?: string;
  email?: string;
  name?: string;
};

type StoredGoogleOAuthCredential = OAuthCredential & {
  clientId?: string;
  clientSecret?: string;
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

function buildGoogleAuthUrl(params: {
  clientId: string;
  state: string;
  challenge: string;
}): string {
  const query = new URLSearchParams({
    client_id: params.clientId,
    response_type: "code",
    redirect_uri: GOOGLE_REDIRECT_URI,
    scope: GOOGLE_SCOPES.join(" "),
    code_challenge: params.challenge,
    code_challenge_method: "S256",
    state: params.state,
    access_type: "offline",
    prompt: "consent",
  });
  return `${GOOGLE_AUTH_URL}?${query.toString()}`;
}

function parseGoogleCallbackInput(
  input: string,
  expectedState: string,
): { code: string; state: string } | { error: string } {
  const trimmed = input.trim();
  if (!trimmed) {
    return { error: "No input provided." };
  }

  try {
    const url = new URL(trimmed);
    const code = url.searchParams.get("code")?.trim();
    const state = url.searchParams.get("state")?.trim();
    const error = url.searchParams.get("error")?.trim();
    if (error) {
      return { error: `Google OAuth error: ${error}` };
    }
    if (!code) {
      return { error: "Missing code parameter. Paste the full redirect URL." };
    }
    if (!state || state !== expectedState) {
      return { error: "OAuth state mismatch. Start login again." };
    }
    return { code, state };
  } catch {
    return { code: trimmed, state: expectedState };
  }
}

async function waitForLocalCallback(params: {
  expectedState: string;
  timeoutMs: number;
  onProgress?: (message: string) => void;
}): Promise<{ code: string; state: string }> {
  const port = 8085;
  const hostname = "localhost";
  const expectedPath = "/oauth2callback";

  return await new Promise<{ code: string; state: string }>((resolve, reject) => {
    let timeout: NodeJS.Timeout | null = null;
    let finished = false;
    const finish = (error?: Error, result?: { code: string; state: string }) => {
      if (finished) {
        return;
      }
      finished = true;
      if (timeout) {
        clearTimeout(timeout);
      }
      try {
        server.close();
      } catch {
        // ignore close failures
      }
      if (error) {
        reject(error);
        return;
      }
      resolve(result ?? { code: "", state: params.expectedState });
    };

    const server = createServer((req, res) => {
      try {
        const requestUrl = new URL(req.url ?? "/", `http://${hostname}:${port}`);
        if (requestUrl.pathname !== expectedPath) {
          res.statusCode = 404;
          res.end("Not found");
          return;
        }

        const error = requestUrl.searchParams.get("error")?.trim();
        const code = requestUrl.searchParams.get("code")?.trim();
        const state = requestUrl.searchParams.get("state")?.trim();

        if (error) {
          res.statusCode = 400;
          res.end(`Authentication failed: ${error}`);
          finish(new Error(`Google OAuth failed: ${error}`));
          return;
        }
        if (!code || !state) {
          res.statusCode = 400;
          res.end("Missing code or state");
          finish(new Error("Missing Google OAuth code or state."));
          return;
        }
        if (state !== params.expectedState) {
          res.statusCode = 400;
          res.end("Invalid state");
          finish(new Error("OAuth state mismatch."));
          return;
        }

        res.statusCode = 200;
        res.setHeader("Content-Type", "text/html; charset=utf-8");
        res.end(
          "<!doctype html><html><body><h2>Executive Assistant Google OAuth complete</h2>" +
            "<p>You can close this window and return to OpenClaw.</p></body></html>",
        );
        finish(undefined, { code, state });
      } catch (error) {
        finish(error instanceof Error ? error : new Error("OAuth callback failed."));
      }
    });

    server.once("error", (error) => {
      finish(error instanceof Error ? error : new Error("OAuth callback server failed."));
    });

    server.listen(port, hostname, () => {
      params.onProgress?.(`Waiting for Google OAuth callback on ${GOOGLE_REDIRECT_URI}…`);
    });

    timeout = setTimeout(() => {
      finish(new Error("Timed out waiting for Google OAuth callback."));
    }, params.timeoutMs);
  });
}

async function exchangeGoogleCodeForTokens(params: {
  clientId: string;
  clientSecret?: string;
  code: string;
  verifier: string;
}): Promise<{ access: string; refresh: string; expires: number }> {
  ensureGlobalUndiciEnvProxyDispatcher();

  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    },
    body: toFormUrlEncoded({
      client_id: params.clientId,
      ...(params.clientSecret ? { client_secret: params.clientSecret } : {}),
      code: params.code,
      grant_type: "authorization_code",
      redirect_uri: GOOGLE_REDIRECT_URI,
      code_verifier: params.verifier,
    }),
  });

  const raw = await response.text();
  if (!response.ok) {
    throw new Error(`Google token exchange failed: ${truncateErrorText(raw) || response.status}`);
  }
  const payload = JSON.parse(raw) as GoogleTokenResponse;
  if (!payload.access_token || !payload.refresh_token || !payload.expires_in) {
    throw new Error(
      payload.error_description?.trim() ||
        "Google token exchange did not return access_token, refresh_token, and expires_in.",
    );
  }

  return {
    access: payload.access_token,
    refresh: payload.refresh_token,
    expires: Date.now() + payload.expires_in * 1000 - 60_000,
  };
}

async function refreshGoogleOAuthCredential(
  credential: StoredGoogleOAuthCredential,
): Promise<StoredGoogleOAuthCredential> {
  const clientId =
    trimOrUndefined(credential.clientId) ??
    trimOrUndefined(process.env[GOOGLE_OAUTH_CLIENT_ID_ENV]);
  if (!clientId) {
    throw new Error(
      `Google OAuth refresh requires clientId on the auth profile or ${GOOGLE_OAUTH_CLIENT_ID_ENV}.`,
    );
  }
  const clientSecret =
    trimOrUndefined(credential.clientSecret) ??
    trimOrUndefined(process.env[GOOGLE_OAUTH_CLIENT_SECRET_ENV]);
  if (!credential.refresh?.trim()) {
    throw new Error("Google OAuth refresh token is missing.");
  }

  ensureGlobalUndiciEnvProxyDispatcher();

  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    },
    body: toFormUrlEncoded({
      client_id: clientId,
      ...(clientSecret ? { client_secret: clientSecret } : {}),
      grant_type: "refresh_token",
      refresh_token: credential.refresh,
    }),
  });

  const raw = await response.text();
  if (!response.ok) {
    throw new Error(`Google token refresh failed: ${truncateErrorText(raw) || response.status}`);
  }
  const payload = JSON.parse(raw) as GoogleTokenResponse;
  if (!payload.access_token || !payload.expires_in) {
    throw new Error(
      payload.error_description?.trim() ||
        "Google token refresh did not return access_token and expires_in.",
    );
  }

  return {
    ...credential,
    type: "oauth",
    provider: GOOGLE_AUTH_PROVIDER_ID,
    access: payload.access_token,
    refresh: payload.refresh_token?.trim() || credential.refresh,
    expires: Date.now() + payload.expires_in * 1000 - 60_000,
    clientId,
    ...(clientSecret ? { clientSecret } : {}),
  };
}

async function fetchGoogleIdentity(accessToken: string): Promise<GoogleUserInfo> {
  ensureGlobalUndiciEnvProxyDispatcher();

  const response = await fetch(GOOGLE_USERINFO_URL, {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
    },
  });
  const raw = await response.text();
  if (!response.ok) {
    throw new Error(`Google userinfo failed: ${truncateErrorText(raw) || response.status}`);
  }
  return JSON.parse(raw) as GoogleUserInfo;
}

async function promptGoogleClientConfig(
  ctx: ProviderAuthContext,
): Promise<{ clientId: string; clientSecret?: string }> {
  const envClientId = trimOrUndefined(process.env[GOOGLE_OAUTH_CLIENT_ID_ENV]);
  const envClientSecret = trimOrUndefined(process.env[GOOGLE_OAUTH_CLIENT_SECRET_ENV]);

  const clientId =
    envClientId ??
    String(
      await ctx.prompter.text({
        message: "Enter Google OAuth client id",
        placeholder: "1234567890-abc123.apps.googleusercontent.com",
        validate: (value) => (value.trim() ? undefined : "Required"),
      }),
    ).trim();

  const clientSecret =
    envClientSecret ??
    trimOrUndefined(
      String(
        await ctx.prompter.text({
          message: "Enter Google OAuth client secret (optional for desktop PKCE clients)",
          placeholder: "Leave blank when your Google OAuth app does not use a secret",
        }),
      ),
    );

  return { clientId, ...(clientSecret ? { clientSecret } : {}) };
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
}): ProviderAuthResult {
  const profileId = buildAuthProfileId({
    providerId: GOOGLE_AUTH_PROVIDER_ID,
    profileName: params.email ?? params.displayName ?? "default",
  });
  const credential: OAuthCredential = {
    type: "oauth",
    provider: GOOGLE_AUTH_PROVIDER_ID,
    access: params.access,
    refresh: params.refresh,
    expires: params.expires,
    ...(params.email ? { email: params.email } : {}),
    ...(params.displayName ? { displayName: params.displayName } : {}),
    ...(params.accountId ? { accountId: params.accountId } : {}),
    clientId: params.clientId,
    ...(params.clientSecret ? { clientSecret: params.clientSecret } : {}),
  } as OAuthCredential;

  return {
    profiles: [{ profileId, credential }],
    configPatch: buildExecutiveAssistantConfigPatch({
      providerId: params.logicalProviderId,
      authProfileId: profileId,
    }),
    notes: [
      `Google account: ${params.email ?? params.displayName ?? "connected"}`,
      "Mail stays read-only. Calendar writes still require writableCalendarIds plus confirm=true.",
      `Legacy direct access-token fallback remains available via ${GOOGLE_ACCESS_TOKEN_ENV}, but authProfileId is the preferred runtime path.`,
    ],
  };
}

export async function runExecutiveAssistantGoogleOAuth(
  ctx: ProviderAuthContext,
): Promise<ProviderAuthResult> {
  const useManualFlow = ctx.isRemote || isWSL2Sync();
  const { clientId, clientSecret } = await promptGoogleClientConfig(ctx);
  const { verifier, challenge } = generatePkceVerifierChallenge();
  const state = randomBytes(16).toString("hex");
  const authUrl = buildGoogleAuthUrl({ clientId, challenge, state });

  await ctx.prompter.note(
    useManualFlow
      ? [
          "Open this URL in your LOCAL browser and sign in to Google.",
          "After Google redirects back, paste the full redirect URL here.",
          "",
          `Redirect URI: ${GOOGLE_REDIRECT_URI}`,
        ].join("\n")
      : [
          "Browser sign-in will open for Google Calendar + Gmail.",
          "If the localhost callback does not complete automatically, paste the final redirect URL back into OpenClaw.",
          "",
          `Redirect URI: ${GOOGLE_REDIRECT_URI}`,
        ].join("\n"),
    "Executive Assistant Google OAuth",
  );

  const progress = ctx.prompter.progress("Starting Google OAuth…");
  try {
    let code: string;
    if (useManualFlow) {
      progress.update("Google OAuth URL ready");
      ctx.runtime.log(`\nOpen this URL in your LOCAL browser:\n\n${authUrl}\n`);
      const pasted = String(
        await ctx.prompter.text({
          message: "Paste the Google redirect URL",
          validate: (value) => (value.trim() ? undefined : "Required"),
        }),
      );
      const parsed = parseGoogleCallbackInput(pasted, state);
      if ("error" in parsed) {
        throw new Error(parsed.error);
      }
      code = parsed.code;
    } else {
      const callback = waitForLocalCallback({
        expectedState: state,
        timeoutMs: 60_000,
        onProgress: (message) => progress.update(message),
      });
      progress.update("Opening browser for Google sign-in…");
      try {
        await ctx.openUrl(authUrl);
      } catch {
        // The URL is still logged below for manual copy/paste.
      }
      ctx.runtime.log(`Open: ${authUrl}`);
      try {
        code = (await callback).code;
      } catch {
        progress.update(
          "Google callback not received automatically. Falling back to manual paste…",
        );
        const pasted = String(
          await ctx.prompter.text({
            message: "Paste the Google redirect URL",
            validate: (value) => (value.trim() ? undefined : "Required"),
          }),
        );
        const parsed = parseGoogleCallbackInput(pasted, state);
        if ("error" in parsed) {
          throw new Error(parsed.error);
        }
        code = parsed.code;
      }
    }

    progress.update("Exchanging Google authorization code…");
    const tokens = await exchangeGoogleCodeForTokens({
      clientId,
      ...(clientSecret ? { clientSecret } : {}),
      code,
      verifier,
    });
    progress.update("Reading Google account identity…");
    const identity = await fetchGoogleIdentity(tokens.access);
    progress.stop("Google OAuth complete");

    return buildAuthResult({
      logicalProviderId: "google",
      ...tokens,
      email: trimOrUndefined(identity.email),
      displayName: trimOrUndefined(identity.name),
      accountId: trimOrUndefined(identity.id),
      clientId,
      ...(clientSecret ? { clientSecret } : {}),
    });
  } catch (error) {
    progress.stop("Google OAuth failed");
    throw error;
  }
}

export function buildExecutiveAssistantGoogleProvider() {
  return {
    id: GOOGLE_AUTH_PROVIDER_ID,
    label: "Executive Assistant Google",
    envVars: [GOOGLE_OAUTH_CLIENT_ID_ENV, GOOGLE_OAUTH_CLIENT_SECRET_ENV],
    auth: [
      {
        id: "oauth",
        label: "Google OAuth",
        hint: "Calendar + Gmail read access with personal-calendar writes gated by policy",
        kind: "oauth" as const,
        wizard: {
          choiceId: "executive-assistant-google",
          choiceLabel: "Executive Assistant Google",
          choiceHint: "Connect Google Calendar + Gmail",
          groupId: "executive-assistant",
          groupLabel: "Executive Assistant",
          groupHint: "Calendar + mail connectors",
        },
        run: async (ctx: ProviderAuthContext) => await runExecutiveAssistantGoogleOAuth(ctx),
      },
    ],
    refreshOAuth: async (credential: OAuthCredential) =>
      await refreshGoogleOAuthCredential(credential as StoredGoogleOAuthCredential),
  };
}
