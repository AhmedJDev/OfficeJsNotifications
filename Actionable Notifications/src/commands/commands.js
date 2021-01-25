/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

let sendEvent;

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true
//   };
//
//   // Show a notification message
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
//
//   // Be sure to indicate when the add-in command function is complete
//   event.completed();
// }

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
// g.action = action;
g.processOnSendEvent = processOnSendEvent;

async function processOnSendEvent(event) {
  sendEvent = event;
  const message1 = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Checking if you are authenticated with Us",
    icon: "Icon.80x80",
    persistent: true
  };

  const message2 = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Checking you are cool enough to use our product",
    icon: "Icon.80x80",
    persistent: true
  };

  const message3 = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: "Something happened! Do not be alarmed but I think you should open the task pane",
		icon: "Icon.80x80",
		actions: [{
    	actionText: 'Open Task Pane',
			actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
			commandId: 'msgComposeOpenPaneButton',
			contextData: '{"SomeObject": "Test"}'
		}]
  };

  // Show a notification message

	Office.context.mailbox.item.notificationMessages.replaceAsync("action", message1);

	await new Promise(resolve => setTimeout(resolve, 3000)); // 3 sec

	Office.context.mailbox.item.notificationMessages.replaceAsync("action", message2);

	await new Promise(resolve => setTimeout(resolve, 3000)); // 3 sec

	Office.context.mailbox.item.notificationMessages.replaceAsync("action", message3);

  // event.completed();
}
