/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 */

// 1. Tell Outlook this function exists
Office.onReady(() => {
    // Ready
});

function validateRecipients(event) {
    // 2. Get recipients
    Office.context.mailbox.item.to.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = result.value;
            
            // NOTE: Background script cannot read "localStorage" from the Taskpane.
            // You must hardcode trusted domains here for the auto-check.
            const trustedDomains = ["paytm.com", "outlook.com"]; 
            
            let externalFound = false;
            let badEmails = [];

            // 3. Scan list
            for (let i = 0; i < recipients.length; i++) {
                let email = recipients[i].emailAddress.toLowerCase();
                let isSafe = false;
                
                for (let j = 0; j < trustedDomains.length; j++) {
                    if (email.indexOf("@" + trustedDomains[j]) > -1) {
                        isSafe = true;
                        break;
                    }
                }
                
                if (!isSafe) {
                    externalFound = true;
                    badEmails.push(email);
                }
            }

            // 4. Block or Allow
            if (externalFound) {
                // BLOCK THE SEND
                console.log("Blocking send.");
                event.completed({
                    allowEvent: false,
                    errorMessage: "⚠️ Security Warning: External recipients detected: " + badEmails.join(", ")
                });
            } else {
                // ALLOW THE SEND
                event.completed({ allowEvent: true });
            }
        } else {
            // If check fails, allow send to avoid getting stuck
            event.completed({ allowEvent: true });
        }
    });
}

// 5. IMPORTANT: Associate the function name from Manifest
Office.actions.associate("validateRecipients", validateRecipients);
