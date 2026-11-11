using System;
using System.Collections.Generic;
using System.Diagnostics;
using ProdigoOutlook.Core;
using ProdigoOutlook.Features.SendLater;
using ProdigoOutlook.OnSend;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.PreSendChecks
{
    /// <summary>
    /// PHASE 2 - Pre-send checks. When an email is sent, reads what the checks need,
    /// runs them (PreSendCheck, no Outlook types), and when anything is found shows one
    /// window listing it all, with Go back (the default) and Send anyway. The rules
    /// are agreed in PreSendCheck's class comment and kept in outlook-rules.json.
    ///
    /// Emails only - meeting invites are left alone, as agreed with Olle.
    ///
    /// Order 100: runs before Send later, so Go back stops the send before anything is
    /// scheduled. Go back is the one legitimate use of SendContext.StopSend. Any error
    /// here means the email is sent: the dispatcher discards a stop request from a
    /// module that threw, and reading failures below simply mean "nothing found".
    ///
    /// An email Prodigo itself puts back in the send queue (Scheduled, OutboxGuard) is
    /// not checked again: Olle decided when he first pressed Send.
    /// </summary>
    internal sealed class PreSendChecksOnSend : IOnSendModule
    {
        public string FeatureId => Feature.PreSendChecks;

        public int Order => 100;

        private const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";
        private const string PrAttachmentHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
        private const string PrAttachContentId = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";

        public void OnSend(object item, SendContext context)
        {
            if (!(item is Outlook.MailItem mail)) return;
            if (SendLaterOutbox.IsResubmitting(SendLaterOnSend.SafeEntryId(mail))) return;

            var timer = Stopwatch.StartNew();
            PreSendRules rules = PreSendRulesFile.Current;
            var input = new PreSendInput
            {
                Body = mail.Body ?? string.Empty,
                RealAttachments = rules.CheckAttachment ? RealAttachments(mail) : 1,
                Recipients = rules.CheckMixed || rules.CheckReplyAll ? Addresses(mail) : new List<string>(),
                IsReply = PreSendCheck.IsReplySubject(mail.Subject),
                Today = DateTime.Today,
            };
            List<PreSendFinding> findings = PreSendCheck.Run(input, rules);
            Log.Info("Pre-send checks: " + findings.Count + " finding(s) for " + Log.Item(SendLaterOnSend.SafeEntryId(mail)) +
                     " in " + timer.ElapsedMilliseconds + " ms" +
                     (findings.Count > 0 ? " (" + string.Join(", ", findings.ConvertAll(f => f.Kind.ToString())) + ")" : ""));
            if (findings.Count == 0) return;

            bool sendAnyway = PreSendWarningForm.Ask(Ol.OwnerWindow(), findings);
            Log.Info("Pre-send checks: user chose " + (sendAnyway ? "Send anyway" : "Go back"));
            if (!sendAnyway) context.StopSend("Go back in pre-send checks");
        }

        /// <summary>
        /// Attachments the reader will see. Pictures inside the text - the signature
        /// logo, or images in the quoted email - are attachments to Outlook too, but
        /// they are hidden or referenced from the text by their content id, and do
        /// not count. When in doubt an attachment counts, so a doubtful case never
        /// produces a warning.
        /// </summary>
        private static int RealAttachments(Outlook.MailItem mail)
        {
            int count = 0;
            string html = null;
            using (var com = new ComScope())
            {
                Outlook.Attachments attachments = com.Track(mail.Attachments);
                int n = attachments.Count;
                for (int i = 1; i <= n; i++)
                {
                    Outlook.Attachment a = com.Track(attachments[i]);
                    try
                    {
                        Outlook.PropertyAccessor pa = com.Track(a.PropertyAccessor);
                        if (TryGet(pa, PrAttachmentHidden) is bool hidden && hidden) continue;
                        if (TryGet(pa, PrAttachContentId) is string cid && cid.Length > 0)
                        {
                            if (html == null) html = SafeHtml(mail);
                            if (html.IndexOf("cid:" + cid, StringComparison.OrdinalIgnoreCase) >= 0) continue;
                        }
                    }
                    catch
                    {
                        // Unreadable: count it.
                    }
                    count++;
                }
            }
            return count;
        }

        /// <summary>Every recipient's SMTP address, also for Exchange users and lists.</summary>
        private static List<string> Addresses(Outlook.MailItem mail)
        {
            var result = new List<string>();
            using (var com = new ComScope())
            {
                Outlook.Recipients recipients = com.Track(mail.Recipients);
                int n = recipients.Count;
                for (int i = 1; i <= n; i++)
                {
                    Outlook.Recipient r = com.Track(recipients[i]);
                    try
                    {
                        string address = TryGet(com.Track(r.PropertyAccessor), PrSmtpAddress) as string;
                        if (string.IsNullOrEmpty(address))
                        {
                            Outlook.AddressEntry entry = com.Track(r.AddressEntry);
                            if (entry != null && string.Equals(entry.Type, "EX", StringComparison.OrdinalIgnoreCase))
                            {
                                Outlook.ExchangeUser user = com.Track(entry.GetExchangeUser());
                                address = user?.PrimarySmtpAddress;
                                if (string.IsNullOrEmpty(address))
                                {
                                    Outlook.ExchangeDistributionList list = com.Track(entry.GetExchangeDistributionList());
                                    address = list?.PrimarySmtpAddress;
                                }
                            }
                            else
                            {
                                address = r.Address;
                            }
                        }
                        if (!string.IsNullOrEmpty(address) && address.Contains("@")) result.Add(address);
                    }
                    catch
                    {
                        // Unreadable recipient: left out, which can only mean fewer warnings.
                    }
                }
            }
            return result;
        }

        private static object TryGet(Outlook.PropertyAccessor pa, string schema)
        {
            try { return pa?.GetProperty(schema); }
            catch { return null; }
        }

        private static string SafeHtml(Outlook.MailItem mail)
        {
            try { return mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML ? mail.HTMLBody ?? string.Empty : string.Empty; }
            catch { return string.Empty; }
        }
    }
}
