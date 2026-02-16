Office.onReady();

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

    // 1. Load Dynamic Trusted Domains
    var savedDomains = Office.context.roamingSettings.get("TrustedDomains");
    var defaultDomains = [
        "paytmpayments.com", "paytmmoney.com", "paytminsurance.co.in", 
        "paytmservices.com", "paytm.com", "powerplay.today", "inapaq.com", 
        "paytmmall.io", "cloud.paytm.com", "firstgames.id", "ticketnew.com", 
        "paytmmall.com", "paytmplay.com", "mobiquest.com", "fellowinfotech.com", 
        "paytminsuretech.com", "alpineinfocom.com", "firstgames.in", "first.games", 
        "paytmfoundation.org", "paytmforbusiness.in", "ps.paytm.com", 
        "paytmcloud.in", "paytm.insure", "mypaytm.com", "paytm.business", 
        "fincollect.in", "creditmate.in", "gamepind.com", "insider.paytm.com", 
        "pmltp.com", "finmate.tech", "cdo.paytm.com", "paytmoffers.in", 
        "paytmmloyal.com", "ocltp.com", "paytm.ca", "quarkinfocom.com", 
        "pibpltp.com", "paytmfirstgames.com", "paytmgic.com", "paytmwholesale.com", 
        "paytmlabs.com", "info.paytmfirstgames.com", "acumengame.com", 
        "robustinfo.com", "one97.sg"
    ];
    var trustedDomains = savedDomains ? JSON.parse(savedDomains) : defaultDomains;

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true }); return;
        }

        var recipients = result.value;
        var externalEmails = [];

        // 2. Check logic
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.indexOf("@" + trustedDomains[j]) > -1) {
                    isSafe = true; break;
                }
            }
            if (!isSafe) { externalEmails.push(email); }
        }

        // 3. Decision Logic
        if (externalEmails.length === 0) {
            // Internal Only -> Send Silently
            event.completed({ allowEvent: true });
        } else {
            // External Found -> Trigger Soft Block
            item.loadCustomPropertiesAsync(function (propResult) {
                var props = propResult.value;
                var warningStatus = props.get("WarningBypass"); 

                if (warningStatus === "yes") {
                    // User clicked Send AGAIN -> Allow it
                    props.remove("WarningBypass");
                    props.saveAsync(function() {
                         event.completed({ allowEvent: true });
                    });
                } else {
                    // FIRST CLICK -> Block and show Native Warning Banner
                    props.set("WarningBypass", "yes");
                    props.saveAsync(function() {
                        event.completed({ 
                            allowEvent: false, 
                            errorMessage: "⚠️ EXTERNALS FOUND: " + externalEmails.join(", ") + ". Click Send again to allow, or use the PayTM Side Panel to remove them." 
                        });
                    });
                }
            });
        }
    });
}
