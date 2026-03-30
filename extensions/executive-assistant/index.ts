import { definePluginEntry, type AnyAgentTool } from "openclaw/plugin-sdk/plugin-entry";
import { createExecutiveAssistantTools } from "./src/tools.js";

export default definePluginEntry({
  id: "executive-assistant",
  name: "Executive Assistant",
  description: "OpenAI-first calendar and mail tools for executive-assistant workflows.",
  register(api) {
    for (const tool of createExecutiveAssistantTools(api)) {
      api.registerTool(tool as AnyAgentTool);
    }
  },
});
