import type { Command } from "commander";
import { defaultRuntime } from "openclaw/plugin-sdk/runtime-env";
import { modelsAuthLoginCommand } from "../../../src/commands/models/auth.js";
import { GOOGLE_AUTH_PROVIDER_ID, MICROSOFT_AUTH_PROVIDER_ID } from "./config.js";

function resolveAuthProviderId(provider: string): string[] {
  switch (provider.trim().toLowerCase()) {
    case "google":
      return [GOOGLE_AUTH_PROVIDER_ID];
    case "microsoft":
      return [MICROSOFT_AUTH_PROVIDER_ID];
    case "all":
      return [GOOGLE_AUTH_PROVIDER_ID, MICROSOFT_AUTH_PROVIDER_ID];
    default:
      throw new Error(`Unknown provider "${provider}". Use google, microsoft, or all.`);
  }
}

export function registerExecutiveAssistantCli(program: Command): void {
  const root = program
    .command("executive-assistant")
    .description("Executive assistant plugin commands");

  root
    .command("auth")
    .argument("<provider>", "google, microsoft, or all")
    .description("Run executive-assistant OAuth onboarding")
    .action(async (provider: string) => {
      for (const providerId of resolveAuthProviderId(provider)) {
        await modelsAuthLoginCommand({ provider: providerId }, defaultRuntime);
      }
    });
}
