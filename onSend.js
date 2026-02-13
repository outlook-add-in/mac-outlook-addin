/*
 * Warning Only (Soft Block) Logic
 */

Office.onReady();

function allowedToSend(event) {
    var item = Office.context.mailbox.item;

    // 1. Get Recipients
    item.to.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var recipients = result.value;
            var trustedDomains = ["paytm.com"]; 
            var externalFound = false;

            // 2. Check for External
            for (var i = 0; i < recipients.length; i++) {
                var email = recipients[i].emailAddress.toLowerCase();
                var isSafe = false;
                for (var j = 0; j < trustedDomains.length; j++) {
                    if (email.indexOf("@" + trustedDomains[j]) > -1) {
                        isSafe = true; break;
                    }
                }
                if (!isSafe) { externalFound = true; }
            }

            if (externalFound) {
                // 3. Check if we already warned the user
                item.loadCustomPropertiesAsync(function (propResult) {
                    var props = propResult.value;
                    var alreadyWarned = props.get("WarningShown");

                    if (alreadyWarned) {
                        // User clicked Send AGAIN -> ALLOW IT
                        event.completed({ allowEvent: true });
                    } else {
                        // First time -> BLOCK IT and show warning
                        props.set("WarningShown", true);
                        props.saveAsync(function(saveResult) {
                            event.completed({ 
                                allowEvent: false, 
                                errorMessage: "⚠️ External recipients found. Click Send again to confirm." 
                            });
                        });
                    }
                });
            } else {
                // Safe -> Allow
                event.completed({ allowEvent: true });
            }
        } else {
            event.completed({ allowEvent: true });
        }
    });
}
