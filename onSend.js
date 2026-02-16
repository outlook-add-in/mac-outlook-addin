Office.onReady();

var dialog;

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

    // Strictly check for paytm.com only
    var trustedDomains = ["paytm.com"];

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var recipients = result.value;
        var externalEmails = []; 

        // 1. Check logic
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.indexOf("@" + trustedDomains[j]) > -1) {
                    isSafe = true; break;
                }
            }
            if (!isSafe) { 
                externalEmails.push(email); 
            }
        }

        // 2. Decision Logic
        if (externalEmails.length === 0) {
            // Internal Only -> Send Silently
            event.completed({ allowEvent: true });
        } else {
            // External Found -> Open simplified popup window
            var encodedEmails = encodeURIComponent(externalEmails.join(","));
            var url = "https://vikash3pandey-sys.github.io/outlook-alerts/warning.html?ext=" + encodedEmails;

            // Open dialog
            Office.context.ui.displayDialogAsync(url, { height: 40, width: 35, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        event.completed({ allowEvent: false, errorMessage: "Security Check Failed." });
                    } else {
                        dialog = asyncResult.value;
                        
                        // Wait for user to click Yes or No
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            dialog.close(); 
                            
                            if (arg.message === "allow") {
                                event.completed({ allowEvent: true }); // Send it
                            } else {
                                event.completed({ allowEvent: false }); // Cancel send
                            }
                        });
                    }
                }
            );
        }
    });
}
