/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});
let dialog;

// The command function.
async function LoginOrLogout(event) {
  // Implement your custom code here. The following code is a simple Word example.
  try {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/dialog.html",
      { height: 30, width: 20 },

      function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    );
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function processMessage(arg) {
  // change font color of selection to be red
  // arg.message is the message from the dialog
  await Word.run(async (context) => {
    // Queue a command to get the current selection.
    // Create a proxy range object for the selection.
    var range = context.document.getSelection();
    if (arg.message == "2") {
      range.insertText("Received message: Logged out\n", "Replace");
    } else {
      range.insertText("Received message: Logged in\n", "Replace");
    }
    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
  });

  OfficeRuntime.storage.setItem("myKey", arg.message);
  if (arg.message == "2") {
    // enable buttons in ribbon
    const button: Control = { id: "TaskpaneButton2", enabled: true };
    const parentGroup: Group = { id: "CommandsGroup1", controls: [button] };
    const parentTab: Tab = { id: "TabPlayground", groups: [parentGroup] };
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab] };
    await Office.ribbon.requestUpdate(ribbonUpdater);
    /*
    await Office.ribbon.requestUpdate({
      tabs: [
        {
          id: "TabPlayground",
          groups: [
            {
              id: "CommandsGroup1",
              controls: [
                {
                  id: "TaskpaneButton2",
                  enabled: true,
                },
              ],
            },
          ],
        },
      ],
    });*/
    await Word.run(async (context) => {
      var range = context.document.getSelection();
      range.insertText("finished logging out\n", "Replace");
      // Synchronize the document state by executing the queued commands,
      // and return a promise to indicate task completion.
      await context.sync();
    });
  } else {
    // disable buttons in ribbon
    await Word.run(async (context) => {
      var range = context.document.getSelection();
      range.insertText("Finished logging in\n", "Replace");
      // Synchronize the document state by executing the queued commands,
      // and return a promise to indicate task completion.
      await context.sync();
    });
    await Office.ribbon.requestUpdate({
      tabs: [
        {
          id: "TabPlayground",
          groups: [
            {
              id: "CommandsGroup1",
              controls: [
                {
                  id: "TaskpaneButton2",
                  enabled: false,
                },
              ],
            },
          ],
        },
      ],
    });
  }
}

// You must register the function with the following line.
Office.actions.associate("LoginOrLogout", LoginOrLogout);
