# Executive Assistant Plugin

OpenClaw-native calendar and mail tools for an OpenAI-backed executive assistant.

What it adds:

- `calendar_list_events`
- `calendar_find_conflicts`
- `calendar_create_personal_event`
- `mail_search_readonly`
- `mail_get_thread`
- `briefing_daily`

Design constraints:

- OpenAI remains the model backend through the existing bundled `openai` provider.
- Mail is read-only.
- Calendar writes are disabled by default and only work on explicit `writableCalendarIds`.
- `calendar_create_personal_event` requires `confirm=true`, so the model has to gather user confirmation before writing.

## Config

Configure the plugin under `plugins.entries.executive-assistant.config`:

```json5
{
  agents: {
    defaults: {
      model: {
        primary: "openai/gpt-5.4-mini",
        fallbacks: ["openai/gpt-5.4"],
      },
      userTimezone: "America/Chicago",
    },
  },
  plugins: {
    entries: {
      "executive-assistant": {
        enabled: true,
        config: {
          defaults: {
            timezone: "America/Chicago",
            lookaheadDays: 3,
            maxCalendarResults: 25,
            maxMailResults: 10,
          },
          google: {
            accessToken: "${OPENCLAW_EXECUTIVE_ASSISTANT_GOOGLE_ACCESS_TOKEN}",
            calendarIds: ["primary"],
            writableCalendarIds: ["primary"],
          },
          microsoft: {
            accessToken: "${OPENCLAW_EXECUTIVE_ASSISTANT_MICROSOFT_ACCESS_TOKEN}",
            userId: "me",
            calendarIds: ["default"],
            writableCalendarIds: [],
          },
        },
      },
    },
  },
}
```

Token env vars:

- `OPENCLAW_EXECUTIVE_ASSISTANT_GOOGLE_ACCESS_TOKEN`
- `OPENCLAW_EXECUTIVE_ASSISTANT_MICROSOFT_ACCESS_TOKEN`

## OpenAI-First Loop

This plugin is meant to run with the existing OpenClaw OpenAI stack, including:

- bundled `openai` provider for `gpt-5.4-mini` / `gpt-5.4`
- existing OpenClaw session loop and tool orchestration
- existing `voice-call` plugin if you want realtime telephony or speech

Example voice-call add-on:

```json5
{
  plugins: {
    entries: {
      "voice-call": {
        enabled: true,
        config: {
          provider: "mock",
          fromNumber: "+15550001234",
          toNumber: "+15550005678",
          outbound: {
            defaultMode: "conversation",
          },
          streaming: {
            enabled: true,
            sttProvider: "openai-realtime",
            sttModel: "gpt-4o-transcribe",
          },
          tts: {
            enabled: true,
            provider: "openai",
            openai: {
              modelId: "gpt-4o-mini-tts",
            },
          },
        },
      },
    },
  },
}
```
