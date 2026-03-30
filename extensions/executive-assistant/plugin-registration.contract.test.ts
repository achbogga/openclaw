import { describePluginRegistrationContract } from "../../test/helpers/plugins/plugin-registration-contract.js";

describePluginRegistrationContract({
  pluginId: "executive-assistant",
  toolNames: [
    "calendar_list_events",
    "calendar_find_conflicts",
    "calendar_create_personal_event",
    "mail_search_readonly",
    "mail_get_thread",
    "briefing_daily",
  ],
});
