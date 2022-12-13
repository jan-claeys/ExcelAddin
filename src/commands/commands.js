/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
import axios from "axios";

const baseUrl = API_URL;

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function publish(args) {
  await Excel.run(async (context) => {
      const newValues = JSON.parse(localStorage.getItem("newValues"));
      const tableId = localStorage.getItem("tableId");
      console.log("publish");
      try {
        res = await axios.post(baseUrl + `/tables/${tableId}`, {
          newValues
        })

        localStorage.setItem('newValues', '');
        download();

      } catch (error) {
        throw error;
      }

      await context.sync();
    })
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  args.completed();
}

function getGlobal() {
  return typeof self !== "undefined" ?
    self :
    typeof window !== "undefined" ?
    window :
    typeof global !== "undefined" ?
    global :
    undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

Office.actions.associate("publish", publish);