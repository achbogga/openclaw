import type { OpenClawPluginApi } from "openclaw/plugin-sdk/plugin-runtime";
import { describe, expect, it, vi } from "vitest";
import { createTestPluginApi } from "../../test/helpers/plugins/plugin-api.js";
import executiveAssistantPlugin from "./index.js";

describe("executive-assistant plugin", () => {
  it("registers the executive assistant toolset", () => {
    const registerTool = vi.fn();
    const api = createTestPluginApi({
      id: "executive-assistant",
      name: "Executive Assistant",
      source: "test",
      config: {},
      runtime: {} as OpenClawPluginApi["runtime"],
      registerTool,
    }) as OpenClawPluginApi;

    executiveAssistantPlugin.register(api);

    const names = registerTool.mock.calls.map((call) => call[0]?.name);
    expect(names).toEqual([
      "calendar_list_events",
      "calendar_find_conflicts",
      "calendar_create_personal_event",
      "mail_search_readonly",
      "mail_get_thread",
      "briefing_daily",
    ]);
  });
});
