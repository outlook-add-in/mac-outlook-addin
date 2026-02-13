/*
 * SOFT BLOCK LOGIC (Warning Only)
 * 1. Internal emails -> Send immediately (No prompt).
 * 2. External emails -> Show Warning once.
 * 3. User clicks Send again -> Email goes through.
 */

Office.onReady();

function allowedToSend(event) {
    var item = Office.context.mailbox.item;

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            // Fail-safe: If we can't read recipients, let it go.
            event.completed({ allowEvent: true });
            return;
        }

        var recipients = result.value;
        // 1. DEFINE YOUR TRUSTED DOMAINS HERE (Lowercase)
        var trustedDomains = ["paytm.com", "paytm.in"]; 
        
        var externalFound = false;
        var externalEmails = [];

        // 2. CHECK RECIPIENTS
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;

            // Check if email ends with any trusted domain
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.indexOf("@" + trustedDomains[j]) > -1) {
                    isSafe = true;
                    break;
                }
            }

            if (!isSafe) {
                externalFound = true;
                externalEmails.push(email);
            }
        }

        // 3. DECISION LOGIC
        if (!externalFound) {
            // CASE A: All recipients are Internal/Safe.
            // ACTION: Send immediately. No prompt.
            event.completed({ allowEvent: true });
        } 
        else {
            // CASE B: External recipients found.
            // Check if we already warned the user for THIS specific email.
            item.loadCustomPropertiesAsync(function (propResult) {
                var props = propResult.value;
                var warningStatus = props.get("WarningShown_V1"); // Unique key

                if (warningStatus === "yes") {
                    // User has already seen the warning and clicked Send AGAIN.
                    // ACTION: Allow the email to send.
                    // (Optional: Clear the flag for next time, though not strictly needed for sent items)
                    props.remove("WarningShown_V1");
                    props.saveAsync(function() {
                         event.completed({ allowEvent: true });
                    });
                } else {
                    // First time user clicked Send.
                    // ACTION: Block email, Show Warning, Set Flag.
                    props.set("WarningShown_V1", "yes");
                    props.saveAsync(function(saveResult) {
                        event.completed({ 
                            allowEvent: false, 
                            errorMessage: "⚠️ External Recipient Warning: You are sending to " + externalEmails.length + " outside address(es). Click Send again to confirm." 
                        });
                    });
                }
            });
        }
    });
}
