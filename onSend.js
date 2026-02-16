Office.onReady();

var dialog;

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

    // 1. Load Dynamic Trusted Domains
    var savedDomains = Office.context.roamingSettings.get("TrustedDomains");
    var trustedDomains = savedDomains ? JSON.parse(savedDomains) : ["paytm.com"];

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var recipients = result.value;
        var externalFound = false;

        // 2. Check logic
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.indexOf("@" + trustedDomains[j]) > -1) {
                    isSafe = true; break;
                }
            }
            if (!isSafe) { externalFound = true; break; }
        }

        if (!externalFound) {
            // Internal Only -> Send Silently
            event.completed({ allowEvent: true });
        } else {
            // External Found -> Open Popup
            var url = "https://vikash3pandey-sys.github.io/outlook-alerts/warning.html";

            Office.context.ui.displayDialogAsync(url, { height: 45, width: 40, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        event.completed({ allowEvent: false, errorMessage: "Security Check Failed." });
                    } else {
                        dialog = asyncResult.value;
                        
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            dialog.close(); 
                            
                            if (arg.message === "allow") {
                                event.completed({ allowEvent: true });
                            } else if (arg.message === "cancel") {
                                event.completed({ allowEvent: false });
                            } else if (arg.message === "remove_and_send") {
                                // USER CHOSE TO CLEAN THE EMAIL
                                removeExternalsAndSend(item, trustedDomains, event);
                            }
                        });
                    }
                }
            );
        }
    });
}

// Helper 1: Removes externals from "To" and "CC" and sends the email
function removeExternalsAndSend(item, trustedDomains, event) {
    // Clean the TO field
    item.to.getAsync(function(resTo) {
        var safeTo = filterSafe(resTo.value, trustedDomains);
        item.to.setAsync(safeTo, function() { 
            
            // Clean the CC field
            item.cc.getAsync(function(resCc) {
                var safeCc = filterSafe(resCc.value, trustedDomains);
                item.cc.setAsync(safeCc, function() {
                    
                    // Fields cleaned -> Allow Send
                    event.completed({ allowEvent: true });
                });
            });
        });
    });
}

// Helper 2: Keeps only trusted emails
function filterSafe(recipients, trustedDomains) {
    if (!recipients) return [];
    var safeList = [];
    
    for (var i = 0; i < recipients.length; i++) {
        var email = recipients[i].emailAddress.toLowerCase();
        var isSafe = false;
        for (var j = 0; j < trustedDomains.length; j++) {
            if (email.indexOf("@" + trustedDomains[j]) > -1) {
                isSafe = true;
                break;
            }
        }
        if (isSafe) {
            safeList.push(recipients[i]);
        }
    }
    return safeList;
}
