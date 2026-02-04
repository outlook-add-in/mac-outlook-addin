/**
 * Logic to check for external recipients
 */
function onMessageSendHandler(event) {
  Office.context.mailbox.item.getToAsync(function (result) {
    const recipients = result.value;
    const externalRecipients = [];

    for (let i = 0; i < recipients.length; i++) {
      let email = recipients[i].emailAddress.toLowerCase();
      if (!email.endsWith("@paytm.com")) {
        externalRecipients.push(email);
      }
    }

    if (externalRecipients.length > 0) {
      // Block the send and show a warning
      event.completed({
        allowEvent: false,
        errorMessage: "EXTERNAL WARNING: You are sending to " + externalRecipients.join(", ") + ". Please verify these addresses are correct before sending."
      });
    } else {
      // Allow the send
      event.completed({ allowEvent: true });
    }
  });
}

// Associate the function name with the manifest
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
