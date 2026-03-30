import { definePluginEntry, type AnyAgentTool } from "openclaw/plugin-sdk/plugin-entry";
import { registerExecutiveAssistantCli } from "./src/cli.js";
import { buildExecutiveAssistantGoogleProvider } from "./src/oauth-google.js";
import { buildExecutiveAssistantMicrosoftProvider } from "./src/oauth-microsoft.js";
import { createExecutiveAssistantTools, EXECUTIVE_ASSISTANT_TOOL_NAMES } from "./src/tools.js";

export default definePluginEntry({
  id: "executive-assistant",
  name: "Executive Assistant",
  description: "OpenAI-first calendar and mail tools for executive-assistant workflows.",
  register(api) {
    api.registerProvider(buildExecutiveAssistantGoogleProvider());
    api.registerProvider(buildExecutiveAssistantMicrosoftProvider());
    api.registerTool(
      (ctx) => {
        const tools = createExecutiveAssistantTools({ api, context: ctx });
        return tools.length > 0 ? (tools as AnyAgentTool[]) : null;
      },
      { names: [...EXECUTIVE_ASSISTANT_TOOL_NAMES] },
    );
    api.registerCli(({ program }) => registerExecutiveAssistantCli(program), {
      commands: ["executive-assistant"],
    });
  },
});
