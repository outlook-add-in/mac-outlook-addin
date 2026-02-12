using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PayTMSafeSend
{
    /// <summary>
    /// PayTM Safe Send - VSTO Add-in for Outlook
    /// Hooks ItemSend event to warn/block external recipients
    /// Works with: Exchange, IMAP, POP3, Gmail in Outlook
    /// </summary>
    public partial class ThisAddIn
    {
        private Outlook.Application outlookApp;
        private Outlook.Inspectors inspectors;
        private List<Outlook.Inspector> openInspectors = new List<Outlook.Inspector>();

        // Configuration
        private SafeSendConfig config = new SafeSendConfig();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                outlookApp = this.Application;
                inspectors = outlookApp.Inspectors;

                // Hook ItemSend event
                outlookApp.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

                // Log startup
                LogEvent("PayTM Safe Send Add-in initialized");
                ShowStartupNotification();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during startup: {ex.Message}", "PayTM Safe Send - Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogEvent($"ERROR: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Main handler - Called before every email send
        /// This is the core functionality
        /// </summary>
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {
                // Only process mail items
                if (!(Item is Outlook.MailItem mailItem))
                    return;

                LogEvent($"ItemSend triggered for: {mailItem.Subject}");

                // Get all recipients
                var recipients = GetAllRecipients(mailItem);

                if (recipients.Count == 0)
                {
                    LogEvent("No recipients found");
                    return;
                }

                // Check for external recipients
                var externalRecipients = CheckRecipients(recipients);

                if (externalRecipients.Count > 0)
                {
                    LogEvent($"External recipients detected: {string.Join(", ", externalRecipients.Select(r => r.Email))}");

                    // Show warning dialog
                    var result = ShowWarningDialog(externalRecipients, mailItem.Subject);

                    if (result == DialogResult.Cancel)
                    {
                        // Block send
                        Cancel = true;
                        LogEvent("Send blocked by user");
                    }
                    else if (result == DialogResult.Yes)
                    {
                        // User confirmed - allow send
                        // Optionally add subject tag
                        if (config.AddSubjectTag && !mailItem.Subject.Contains("[External]"))
                        {
                            mailItem.Subject = $"[External] {mailItem.Subject}";
                            LogEvent($"Subject tagged: {mailItem.Subject}");
                        }

                        // Add footer if configured
                        if (config.AddWarningFooter)
                        {
                            AddWarningFooter(mailItem);
                        }

                        LogEvent("Send allowed with confirmation");
                    }
                }
                else
                {
                    LogEvent("All recipients verified - safe domain only");
                }
            }
            catch (Exception ex)
            {
                LogEvent($"ERROR in ItemSend: {ex.Message}\n{ex.StackTrace}");
                // Don't block send on error - just log
            }
        }

        /// <summary>
        /// Get all recipients from mail item (To, Cc, Bcc)
        /// </summary>
        private List<RecipientInfo> GetAllRecipients(Outlook.MailItem mailItem)
        {
            var recipients = new List<RecipientInfo>();

            try
            {
                // Get To recipients
                if (mailItem.To != null && mailItem.To.Length > 0)
                {
                    var toAddresses = ParseRecipients(mailItem.To);
                    recipients.AddRange(toAddresses);
                }

                // Get Cc recipients
                if (mailItem.CC != null && mailItem.CC.Length > 0)
                {
                    var ccAddresses = ParseRecipients(mailItem.CC);
                    recipients.AddRange(ccAddresses);
                }

                // Get Bcc recipients
                if (mailItem.BCC != null && mailItem.BCC.Length > 0)
                {
                    var bccAddresses = ParseRecipients(mailItem.BCC);
                    recipients.AddRange(bccAddresses);
                }

                LogEvent($"Total recipients: {recipients.Count}");
            }
            catch (Exception ex)
            {
                LogEvent($"Error getting recipients: {ex.Message}");
            }

            return recipients;
        }

        /// <summary>
        /// Parse recipient addresses from string
        /// </summary>
        private List<RecipientInfo> ParseRecipients(string recipientString)
        {
            var recipients = new List<RecipientInfo>();

            if (string.IsNullOrEmpty(recipientString))
                return recipients;

            // Split by semicolon (Outlook separator)
            var addresses = recipientString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var address in addresses)
            {
                var trimmedAddress = address.Trim();
                if (!string.IsNullOrEmpty(trimmedAddress))
                {
                    recipients.Add(new RecipientInfo
                    {
                        Email = ExtractEmailAddress(trimmedAddress),
                        DisplayName = ExtractDisplayName(trimmedAddress)
                    });
                }
            }

            return recipients;
        }

        /// <summary>
        /// Extract email address from recipient
        /// Handles formats: email@domain.com, Name <email@domain.com>, "Name" <email@domain.com>
        /// </summary>
        private string ExtractEmailAddress(string recipient)
        {
            recipient = recipient.Trim();

            // Format: Name <email@domain.com>
            if (recipient.Contains("<") && recipient.Contains(">"))
            {
                int start = recipient.LastIndexOf('<');
                int end = recipient.LastIndexOf('>');
                if (start >= 0 && end > start)
                {
                    return recipient.Substring(start + 1, end - start - 1).Trim();
                }
            }

            // Format: email@domain.com
            if (recipient.Contains("@"))
            {
                return recipient;
            }

            return recipient;
        }

        /// <summary>
        /// Extract display name from recipient
        /// </summary>
        private string ExtractDisplayName(string recipient)
        {
            if (recipient.Contains("<") && recipient.Contains(">"))
            {
                int start = recipient.IndexOf('<');
                return recipient.Substring(0, start).Trim().Trim('"');
            }
            return recipient;
        }

        /// <summary>
        /// Check which recipients are external (not in trusted domains)
        /// </summary>
        private List<RecipientInfo> CheckRecipients(List<RecipientInfo> recipients)
        {
            var external = new List<RecipientInfo>();

            foreach (var recipient in recipients)
            {
                string domain = ExtractDomain(recipient.Email);

                // Check if domain is trusted
                bool isTrusted = config.TrustedDomains.Contains(domain, StringComparer.OrdinalIgnoreCase) ||
                                config.AllowList.Contains(domain, StringComparer.OrdinalIgnoreCase);

                if (!isTrusted && !string.IsNullOrEmpty(domain))
                {
                    recipient.Domain = domain;
                    external.Add(recipient);
                }
            }

            return external;
        }

        /// <summary>
        /// Extract domain from email address
        /// </summary>
        private string ExtractDomain(string email)
        {
            if (string.IsNullOrEmpty(email) || !email.Contains("@"))
                return string.Empty;

            return email.Substring(email.LastIndexOf('@') + 1).ToLower();
        }

        /// <summary>
        /// Show warning dialog to user
        /// Returns: Yes = Send, No = Send and don't warn again, Cancel = Block
        /// </summary>
        private DialogResult ShowWarningDialog(List<RecipientInfo> externalRecipients, string subject)
        {
            string recipientList = string.Join("\n", 
                externalRecipients.Select(r => $"  • {r.Email} ({r.Domain})"));

            string message = $@"⚠️ WARNING: External Recipients Detected

Subject: {subject}

You are sending to recipients outside the paytm.com domain:

{recipientList}

Only paytm.com addresses are trusted. 
Be careful to verify these recipients are intended.

Do you want to proceed?";

            var result = MessageBox.Show(message,
                "PayTM Safe Send - Confirm Send",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button3); // Default to Cancel

            return result;
        }

        /// <summary>
        /// Add warning footer to email body
        /// </summary>
        private void AddWarningFooter(Outlook.MailItem mailItem)
        {
            try
            {
                string footer = "\n\n---\n⚠️ This email was sent to external (non-PayTM) recipients.";

                if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatPlain)
                {
                    mailItem.Body += footer;
                }
                else if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                {
                    mailItem.HTMLBody += $"<p style='color: #d32f2f; font-size: 11px;'>⚠️ This email was sent to external (non-PayTM) recipients.</p>";
                }

                LogEvent("Warning footer added to email");
            }
            catch (Exception ex)
            {
                LogEvent($"Error adding footer: {ex.Message}");
            }
        }

        /// <summary>
        /// Show startup notification
        /// </summary>
        private void ShowStartupNotification()
        {
            MessageBox.Show(
                "PayTM Safe Send is now active.\n\n" +
                "You will receive a warning when sending to external (non-PayTM) recipients.\n\n" +
                "Trusted domains: paytm.com",
                "PayTM Safe Send",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// Log event to file for debugging
        /// </summary>
        private void LogEvent(string message)
        {
            try
            {
                string logPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "PayTMSafeSend",
                    "log.txt");

                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(logPath));

                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
                System.IO.File.AppendAllText(logPath, logMessage + Environment.NewLine);
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            LogEvent("PayTM Safe Send shut down");
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }

    /// <summary>
    /// Recipient information structure
    /// </summary>
    public class RecipientInfo
    {
        public string Email { get; set; }
        public string DisplayName { get; set; }
        public string Domain { get; set; }
    }

    /// <summary>
    /// Configuration for SafeSend
    /// </summary>
    public class SafeSendConfig
    {
        public List<string> TrustedDomains { get; set; } = new List<string> { "paytm.com" };
        
        public List<string> AllowList { get; set; } = new List<string>
        {
            // Add partner domains here
            // "partner.com",
            // "vendor.com"
        };

        public bool BlockMode { get; set; } = true;
        public bool AddSubjectTag { get; set; } = true;
        public bool AddWarningFooter { get; set; } = false;
        public bool LogEnabled { get; set; } = true;
    }
}
