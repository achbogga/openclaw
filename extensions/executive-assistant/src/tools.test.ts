import type { OpenClawPluginApi } from "openclaw/plugin-sdk/plugin-runtime";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { createTestPluginApi } from "../../../test/helpers/plugins/plugin-api.js";
import { createExecutiveAssistantTools } from "./tools.js";

const mocks = vi.hoisted(() => ({
  loadAuthProfileStoreForRuntime: vi.fn(() => ({ profiles: {} })),
  resolveApiKeyForProfile: vi.fn(),
}));

vi.mock("openclaw/plugin-sdk/agent-runtime", () => ({
  loadAuthProfileStoreForRuntime: mocks.loadAuthProfileStoreForRuntime,
  resolveApiKeyForProfile: mocks.resolveApiKeyForProfile,
}));

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function createApi(config?: Record<string, unknown>) {
  return createTestPluginApi({
    id: "executive-assistant",
    name: "Executive Assistant",
    source: "test",
    config: config ?? {},
    runtime: {} as OpenClawPluginApi["runtime"],
  }) as OpenClawPluginApi;
}

function findTool(
  name: string,
  config?: Record<string, unknown>,
  context?: { agentDir?: string; runtimeConfig?: Record<string, unknown> },
) {
  const tools = createExecutiveAssistantTools({ api: createApi(config), context });
  const tool = tools.find((entry) => entry.name === name);
  if (!tool) {
    throw new Error(`tool ${name} not found`);
  }
  return tool;
}

describe("executive-assistant tools", () => {
  const fetchMock = vi.fn<typeof fetch>();

  const config = {
    plugins: {
      entries: {
        "executive-assistant": {
          config: {
            google: {
              accessToken: "google-token",
              calendarIds: ["primary"],
              writableCalendarIds: ["primary"],
            },
            microsoft: {
              accessToken: "microsoft-token",
              calendarIds: ["default"],
            },
          },
        },
      },
    },
  };

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal("fetch", fetchMock);
    mocks.loadAuthProfileStoreForRuntime.mockReset();
    mocks.loadAuthProfileStoreForRuntime.mockReturnValue({ profiles: {} });
    mocks.resolveApiKeyForProfile.mockReset();
  });

  it("lists calendar events across Google and Microsoft providers", async () => {
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          items: [
            {
              id: "g-1",
              summary: "Personal sync",
              start: { dateTime: "2026-03-31T14:00:00Z" },
              end: { dateTime: "2026-03-31T14:30:00Z" },
            },
          ],
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({
          value: [
            {
              id: "m-1",
              subject: "Design review",
              start: { dateTime: "2026-03-31T15:00:00", timeZone: "UTC" },
              end: { dateTime: "2026-03-31T16:00:00", timeZone: "UTC" },
            },
          ],
        }),
      );

    const tool = findTool("calendar_list_events", config);
    const result = await tool.execute("call-1", {
      start_time: "2026-03-31T00:00:00Z",
      end_time: "2026-04-01T00:00:00Z",
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(result.details).toMatchObject({
      events: [
        expect.objectContaining({ provider: "google", id: "g-1", title: "Personal sync" }),
        expect.objectContaining({ provider: "microsoft", id: "m-1", title: "Design review" }),
      ],
    });
  });

  it("requires explicit confirmation before calendar writes", async () => {
    const tool = findTool("calendar_create_personal_event", config);
    await expect(
      tool.execute("call-2", {
        provider: "google",
        title: "Dinner",
        start_time: "2026-03-31T18:00:00Z",
        end_time: "2026-03-31T19:00:00Z",
      }),
    ).rejects.toThrow(/confirm=true/);
  });

  it("creates an allowlisted Google calendar event after confirmation", async () => {
    fetchMock.mockResolvedValueOnce(
      jsonResponse({
        id: "created-1",
        summary: "Dinner",
        start: { dateTime: "2026-03-31T18:00:00Z" },
        end: { dateTime: "2026-03-31T19:00:00Z" },
      }),
    );

    const tool = findTool("calendar_create_personal_event", config);
    const result = await tool.execute("call-3", {
      provider: "google",
      calendar_id: "primary",
      title: "Dinner",
      start_time: "2026-03-31T18:00:00Z",
      end_time: "2026-03-31T19:00:00Z",
      confirm: true,
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("/calendars/primary/events");
    expect(result.details).toMatchObject({
      confirmed: true,
      event: {
        provider: "google",
        id: "created-1",
        title: "Dinner",
      },
    });
  });

  it("searches read-only mail across both providers", async () => {
    fetchMock.mockImplementation(async (input) => {
      const url = String(input);
      if (url.includes("/gmail/v1/users/me/messages?")) {
        return jsonResponse({
          messages: [{ id: "gm-1", threadId: "gt-1" }],
        });
      }
      if (url.includes("/gmail/v1/users/me/messages/gm-1")) {
        return jsonResponse({
          id: "gm-1",
          threadId: "gt-1",
          internalDate: String(Date.parse("2026-03-31T12:00:00Z")),
          snippet: "Plan for tomorrow",
          labelIds: ["UNREAD"],
          payload: {
            headers: [
              { name: "Subject", value: "Plan" },
              { name: "From", value: "ceo@example.com" },
              { name: "Date", value: "Tue, 31 Mar 2026 12:00:00 +0000" },
            ],
          },
        });
      }
      if (url.includes("graph.microsoft.com")) {
        return jsonResponse({
          value: [
            {
              id: "mm-1",
              conversationId: "mt-1",
              subject: "Plan draft",
              from: { emailAddress: { address: "boss@example.com", name: "Boss" } },
              receivedDateTime: "2026-03-31T13:00:00Z",
              isRead: false,
              bodyPreview: "Need your review",
            },
          ],
        });
      }
      throw new Error(`Unexpected fetch URL: ${url}`);
    });

    const tool = findTool("mail_search_readonly", config);
    const result = await tool.execute("call-4", {
      query: "plan",
    });

    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(result.details).toMatchObject({
      messages: [
        expect.objectContaining({ provider: "microsoft", id: "mm-1", threadId: "mt-1" }),
        expect.objectContaining({ provider: "google", id: "gm-1", threadId: "gt-1" }),
      ],
    });
  });

  it("resolves auth-profile-backed tokens through the OpenClaw auth store", async () => {
    mocks.resolveApiKeyForProfile.mockResolvedValue({
      apiKey: "google-profile-token",
      provider: "executive-assistant-google",
      email: "assistant@example.com",
    });
    fetchMock.mockResolvedValueOnce(
      jsonResponse({
        items: [
          {
            id: "g-auth-1",
            summary: "Profile-backed sync",
            start: { dateTime: "2026-03-31T14:00:00Z" },
            end: { dateTime: "2026-03-31T14:30:00Z" },
          },
        ],
      }),
    );

    const tool = findTool(
      "calendar_list_events",
      {
        plugins: {
          entries: {
            "executive-assistant": {
              config: {
                google: {
                  authProfileId: "executive-assistant-google:assistant@example.com",
                  calendarIds: ["primary"],
                },
              },
            },
          },
        },
        auth: {
          profiles: {
            "executive-assistant-google:assistant@example.com": {
              provider: "executive-assistant-google",
              mode: "oauth",
            },
          },
        },
      },
      { agentDir: "/tmp/openclaw-agent" },
    );

    const result = await tool.execute("call-5", {
      provider: "google",
      start_time: "2026-03-31T00:00:00Z",
      end_time: "2026-04-01T00:00:00Z",
    });

    expect(mocks.loadAuthProfileStoreForRuntime).toHaveBeenCalledWith("/tmp/openclaw-agent");
    expect(mocks.resolveApiKeyForProfile).toHaveBeenCalledWith(
      expect.objectContaining({
        profileId: "executive-assistant-google:assistant@example.com",
        agentDir: "/tmp/openclaw-agent",
      }),
    );
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0]?.[1]).toMatchObject({
      headers: expect.objectContaining({
        Authorization: "Bearer google-profile-token",
      }),
    });
    expect(result.details).toMatchObject({
      events: [expect.objectContaining({ provider: "google", id: "g-auth-1" })],
    });
  });
});
