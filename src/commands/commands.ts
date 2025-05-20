import { setNotificationInOutlook } from "./outlook";

/* global Office */

// Register the add-in commands with the Office host application.
Office.onReady(async (info) => {
  switch (info.host) {
    case Office.HostType.Outlook:
      Office.actions.associate("action", setNotificationInOutlook);
      break;
    default: {
      throw new Error(`${info.host} not supported.`);
    }
  }
});
