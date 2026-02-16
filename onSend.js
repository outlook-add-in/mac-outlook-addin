Office.onReady();

var dialog;

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

    // Your complete list of trusted domains
    var trustedDomains = [
        "paytmpayments.com", "paytmmoney.com", "paytminsurance.co.in", "paytmservices.com", "paytm.com", 
        "powerplay.today", "inapaq.com", "paytmmall.io", "cloud.paytm.com", "firstgames.id", "ticketnew.com", 
        "paytmmall.com", "paytmplay.com", "mobiquest.com", "fellowinfotech.com", "paytminsuretech.com", 
        "alpineinfocom.com", "firstgames.in", "first.games", "paytmfoundation.org", "paytmforbusiness.in", 
        "ps.paytm.com", "paytmcloud.in", "paytm.insure", "mypaytm.com", "paytm.business", "fincollect.in", 
        "creditmate.in", "gamepind.com", "insider.paytm.com", "pmltp.com", "finmate.tech", "cdo.paytm.com", 
        "paytmoffers.in", "paytmmloyal.com", "ocltp.com", "paytm.ca", "quarkinfocom.com", "pibpltp.com", 
        "paytmfirstgames.com", "paytmgic.com", "paytmwholesale.com", "paytmlabs.com", "info.paytmfirstgames.com", 
        "acumengame.com", "robustinfo.com", "one97.sg"
    ];

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var recipients = result.value;
        var externalEmails = []; 

        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.endsWith("@" + trustedDomains[j]) || email.endsWith("." + trustedDomains[j])) {
                    isSafe = true; 
                    break;
                }
            }
            if (!isSafe) { 
                externalEmails.push(email); 
            }
        }

        if (externalEmails.length === 0) {
            event.completed({ allowEvent: true });
        } else {
            // UPDATED URL HERE
            var encodedEmails = encodeURIComponent(externalEmails.join(","));
            var url = "https://outlook-add-in.github.io/mac-outlook-addin/warning.html?ext=" + encodedEmails;

            Office.context.ui.displayDialogAsync(url, { height: 40, width: 35, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        event.completed({ allowEvent: false, errorMessage: "Security Check Failed." });
                    } else {
                        dialog = asyncResult.value;
                        
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            dialog.close(); 
                            
                            if (arg.message === "allow") {
                                event.completed({ allowEvent: true }); 
                            } else {
                                event.completed({ allowEvent: false }); 
                            }
                        });
                    }
                }
            );
        }
    });
}
