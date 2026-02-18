Office.onReady();

var dialog;

function checkExternalRecipients(event) {
    try {
        var item = Office.context.mailbox.item;

        // ✅ FIXED: Validate that item exists
        if (!item) {
            console.error("Mailbox item not found");
            event.completed({ allowEvent: false, errorMessage: "Error: Could not access email item" });
            return;
        }

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

        // ✅ FIXED: Validate email address format
        function isValidEmail(email) {
            if (!email || typeof email !== 'string') return false;
            // Simple email validation
            var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            return emailRegex.test(email);
        }

        // ✅ FIXED: Sanitize and validate email
        function sanitizeEmail(email) {
            if (!email || typeof email !== 'string') return '';
            return email.toLowerCase().trim();
        }

        // ✅ FIXED: Check if email is safe based on trusted domains
        function isSafeEmail(email, domains) {
            try {
                var sanitized = sanitizeEmail(email);
                
                if (!isValidEmail(sanitized)) {
                    console.warn("Invalid email format:", email);
                    return false; // Treat invalid emails as external
                }

                // Check against trusted domains
                for (var j = 0; j < domains.length; j++) {
                    var domain = sanitizeEmail(domains[j]);
                    if (!domain) continue;
                    
                    // ✅ More robust domain checking
                    if (sanitized.endsWith("@" + domain) || sanitized.endsWith("." + domain)) {
                        return true;
                    }
                }
                
                return false;
            } catch (error) {
                console.error("Error checking email safety:", error);
                return false;
            }
        }

        // ✅ FIXED: Validate and process recipients
        function processRecipients(recipients) {
            try {
                var externalEmails = [];
                
                // ✅ Validate recipients is an array
                if (!Array.isArray(recipients)) {
                    console.warn("Recipients is not an array:", recipients);
                    recipients = [];
                }

                for (var i = 0; i < recipients.length; i++) {
                    try {
                        var recipient = recipients[i];
                        
                        // ✅ Validate recipient object
                        if (!recipient || !recipient.emailAddress) {
                            console.warn("Invalid recipient object at index:", i);
                            continue;
                        }

                        var email = sanitizeEmail(recipient.emailAddress);
                        
                        // ✅ Skip empty emails
                        if (!email) {
                            console.warn("Empty email at index:", i);
                            continue;
                        }

                        // ✅ Check if email is safe
                        if (!isSafeEmail(email, trustedDomains)) {
                            externalEmails.push(email);
                        }
                    } catch (recipientError) {
                        console.error("Error processing recipient:", recipientError);
                        continue;
                    }
                }

                // ✅ Remove duplicates from external emails
                externalEmails = Array.from(new Set(externalEmails));

                // No external emails - safe to send
                if (externalEmails.length === 0) {
                    event.completed({ allowEvent: true });
                } else {
                    // External emails found - show warning dialog
                    showWarningDialog(externalEmails, event);
                }
            } catch (error) {
                console.error("Error processing recipients:", error);
                event.completed({ allowEvent: false, errorMessage: "Error checking recipients" });
            }
        }

        // ✅ FIXED: Safe dialog display with error handling
        function showWarningDialog(externalEmails, event) {
            try {
                // ✅ Validate email list
                if (!Array.isArray(externalEmails) || externalEmails.length === 0) {
                    console.error("Invalid external emails list");
                    event.completed({ allowEvent: false, errorMessage: "Error: Invalid recipient list" });
                    return;
                }

                // ✅ Properly encode emails for URL
                var encodedEmails = encodeURIComponent(externalEmails.join(","));
                var url = "https://outlook-add-in.github.io/mac-outlook-addin/warning.html?ext=" + encodedEmails;

                // ✅ Validate URL length (avoid URL length limits)
                if (url.length > 2000) {
                    console.warn("URL too long, truncating email list");
                    // Send only first few emails in URL
                    var truncatedEmails = externalEmails.slice(0, 10);
                    encodedEmails = encodeURIComponent(truncatedEmails.join(","));
                    url = "https://outlook-add-in.github.io/mac-outlook-addin/warning.html?ext=" + encodedEmails;
                }

                // ✅ Display dialog with error handling
                Office.context.ui.displayDialogAsync(url, 
                    { height: 40, width: 35, displayInIframe: true },
                    function (asyncResult) {
                        try {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                console.error("Dialog display failed:", asyncResult.error);
                                event.completed({ 
                                    allowEvent: false, 
                                    errorMessage: "Security check failed: " + asyncResult.error.message 
                                });
                            } else {
                                dialog = asyncResult.value;
                                
                                // ✅ Add error handling for dialog message
                                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                                    try {
                                        if (dialog) {
                                            dialog.close();
                                        }
                                        
                                        var message = arg.message;
                                        
                                        // ✅ Validate message
                                        if (message === "allow") {
                                            event.completed({ allowEvent: true });
                                        } else if (message === "remove_and_send") {
                                            // User wants to remove external recipients and send
                                            // This would require additional implementation to remove recipients
                                            event.completed({ allowEvent: false });
                                        } else {
                                            // Default to cancel for any other message
                                            event.completed({ allowEvent: false });
                                        }
                                    } catch (messageError) {
                                        console.error("Error processing dialog message:", messageError);
                                        event.completed({ allowEvent: false, errorMessage: "Error processing response" });
                                    }
                                });

                                // ✅ Add error handler for dialog
                                dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                                    try {
                                        if (arg.error === Office.EventType.DialogClosed || arg.error === 12002) {
                                            // Dialog was closed without sending a message
                                            console.log("Dialog closed by user");
                                            event.completed({ allowEvent: false });
                                        }
                                    } catch (error) {
                                        console.error("Error in dialog event handler:", error);
                                    }
                                });
                            }
                        } catch (asyncError) {
                            console.error("Error in displayDialogAsync callback:", asyncError);
                            event.completed({ allowEvent: false, errorMessage: "Dialog error" });
                        }
                    }
                );
            } catch (error) {
                console.error("Error showing warning dialog:", error);
                event.completed({ allowEvent: false, errorMessage: "Error displaying security dialog" });
            }
        }

        // ✅ FIXED: Determine if item is calendar invite or email with error handling
        function getAndProcessRecipients() {
            try {
                // ✅ Validate item type
                if (!item.itemType) {
                    console.warn("Item type not available");
                    event.completed({ allowEvent: true });
                    return;
                }

                if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
                    // IT IS A CALENDAR INVITE: Check required and optional attendees
                    getRequiredAttendees();
                } else {
                    // IT IS AN EMAIL: Check TO and CC
                    getToRecipients();
                }
            } catch (error) {
                console.error("Error determining item type:", error);
                event.completed({ allowEvent: false, errorMessage: "Error: Could not determine message type" });
            }
        }

        // ✅ FIXED: Safe attendee gathering
        function getRequiredAttendees() {
            try {
                item.requiredAttendees.getAsync(function(reqResult) {
                    try {
                        var allAttendees = [];
                        
                        if (reqResult.status === Office.AsyncResultStatus.Succeeded && reqResult.value) {
                            allAttendees = reqResult.value;
                        } else {
                            console.warn("Could not get required attendees");
                        }

                        getOptionalAttendees(allAttendees);
                    } catch (error) {
                        console.error("Error in getRequiredAttendees callback:", error);
                        event.completed({ allowEvent: false, errorMessage: "Error reading attendees" });
                    }
                });
            } catch (error) {
                console.error("Error getting required attendees:", error);
                event.completed({ allowEvent: false, errorMessage: "Error reading required attendees" });
            }
        }

        // ✅ FIXED: Safe optional attendee gathering
        function getOptionalAttendees(attendees) {
            try {
                item.optionalAttendees.getAsync(function(optResult) {
                    try {
                        var allAttendees = attendees || [];
                        
                        if (optResult.status === Office.AsyncResultStatus.Succeeded && optResult.value) {
                            allAttendees = allAttendees.concat(optResult.value);
                        }

                        processRecipients(allAttendees);
                    } catch (error) {
                        console.error("Error in getOptionalAttendees callback:", error);
                        event.completed({ allowEvent: false, errorMessage: "Error reading optional attendees" });
                    }
                });
            } catch (error) {
                console.error("Error getting optional attendees:", error);
                event.completed({ allowEvent: false, errorMessage: "Error reading optional attendees" });
            }
        }

        // ✅ FIXED: Safe recipient gathering for emails
        function getToRecipients() {
            try {
                item.to.getAsync(function(toResult) {
                    try {
                        var allRecipients = [];
                        
                        if (toResult.status === Office.AsyncResultStatus.Succeeded && toResult.value) {
                            allRecipients = toResult.value;
                        } else {
                            console.warn("Could not get TO recipients");
                        }

                        getCcRecipients(allRecipients);
                    } catch (error) {
                        console.error("Error in getToRecipients callback:", error);
                        event.completed({ allowEvent: false, errorMessage: "Error reading recipients" });
                    }
                });
            } catch (error) {
                console.error("Error getting TO recipients:", error);
                event.completed({ allowEvent: false, errorMessage: "Error reading TO line" });
            }
        }

        // ✅ FIXED: Safe CC recipient gathering
        function getCcRecipients(recipients) {
            try {
                item.cc.getAsync(function(ccResult) {
                    try {
                        var allRecipients = recipients || [];
                        
                        if (ccResult.status === Office.AsyncResultStatus.Succeeded && ccResult.value) {
                            allRecipients = allRecipients.concat(ccResult.value);
                        }

                        processRecipients(allRecipients);
                    } catch (error) {
                        console.error("Error in getCcRecipients callback:", error);
                        event.completed({ allowEvent: false, errorMessage: "Error reading CC line" });
                    }
                });
            } catch (error) {
                console.error("Error getting CC recipients:", error);
                event.completed({ allowEvent: false, errorMessage: "Error reading CC recipients" });
            }
        }

        // ✅ FIXED: Start the process
        getAndProcessRecipients();

    } catch (error) {
        console.error("Error in checkExternalRecipients:", error);
        event.completed({ allowEvent: false, errorMessage: "Security check error: " + error.message });
    }
}
