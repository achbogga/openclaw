import type { OpenClawPluginApi } from "openclaw/plugin-sdk/plugin-runtime";
import { describe, expect, it, vi } from "vitest";
import { createTestPluginApi } from "../../test/helpers/plugins/plugin-api.js";
import executiveAssistantPlugin from "./index.js";
import { EXECUTIVE_ASSISTANT_TOOL_NAMES } from "./src/tools.js";

describe("executive-assistant plugin", () => {
  it("registers assistant providers, tool factory metadata, and CLI", () => {
    const registerTool = vi.fn();
    const registerProvider = vi.fn();
    const registerCli = vi.fn();
    const api = createTestPluginApi({
      id: "executive-assistant",
      name: "Executive Assistant",
      source: "test",
      config: {},
      runtime: {} as OpenClawPluginApi["runtime"],
      registerTool,
      registerProvider,
      registerCli,
    }) as OpenClawPluginApi;

    executiveAssistantPlugin.register(api);

    expect(registerProvider.mock.calls.map((call) => call[0]?.id)).toEqual([
      "executive-assistant-google",
      "executive-assistant-microsoft",
    ]);
    expect(registerTool).toHaveBeenCalledTimes(1);
    expect(typeof registerTool.mock.calls[0]?.[0]).toBe("function");
    expect(registerTool.mock.calls[0]?.[1]).toMatchObject({
      names: [...EXECUTIVE_ASSISTANT_TOOL_NAMES],
    });
    expect(registerCli).toHaveBeenCalledTimes(1);
  });
});
