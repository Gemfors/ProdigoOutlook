using ProdigoOutlook.Core;
using ProdigoOutlook.OnSend;

namespace ProdigoOutlook.Features.PreSendChecks
{
    /// <summary>
    /// PHASE 2 - planned, not built. Warns before an email goes, with "Send anyway" and
    /// "Go back".
    ///
    /// Checks: attachment mentioned but none attached (bifogar, bifogad, bilaga,
    /// attached, attachment, enclosed - ignoring quoted text from earlier in the
    /// thread); empty subject; reply-all to more than N recipients; internal and
    /// external recipients mixed; sending from an unexpected account; weekday and date
    /// that do not match ("tisdag 12 nov" when the 12th is a Wednesday), Swedish and
    /// English.
    ///
    /// Every keyword, threshold and domain lives in %APPDATA%\Prodigo\outlook-rules.json,
    /// not in code, so a future Office.js add-in for new Outlook can read the same file.
    /// The checks themselves belong in a class with no Outlook types (text in, findings
    /// out), so they can be tested in test\LogicTests like SendLaterTime.
    ///
    /// Order 100: runs before Send later, so "Go back" stops the send before anything
    /// is scheduled. "Go back" is the one legitimate use of SendContext.StopSend; any
    /// exception still means the email is sent.
    /// </summary>
    internal sealed class PreSendChecksOnSend : IOnSendModule
    {
        public string FeatureId => Feature.PreSendChecks;

        public int Order => 100;

        public void OnSend(object item, SendContext context)
        {
            // Not built. Feature.PreSendChecks has Built = false, so this is never called.
        }
    }
}
