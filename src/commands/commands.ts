/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
var mailboxItem;
Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
}
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});


function helloOnSend(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Hello world!",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("helloOnSend", message);
  console.log('test');

  mailboxItem.subject.getAsync({asyncContext: event},
    function(asyncResult){
      let subject = asyncResult.value +" [Hello Outlook]";
      mailboxItem.subject.setAsync(
        subject, {asyncContext: asyncResult.asyncContext}, 
        function(asyncResult){
          asyncResult.asyncContext.completed({ allowEvent: true });
        }
      )
    })
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
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

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
g.helloOnSend = helloOnSend;
