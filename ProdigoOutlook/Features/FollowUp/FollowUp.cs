using ProdigoOutlook.Core;
using ProdigoOutlook.OnSend;

namespace ProdigoOutlook.Features.FollowUp
{
    /// <summary>
    /// PHASE 5 - planned, not built. "Remind me if no reply in X days".
    ///
    /// At send time the compose window's Remind If No Reply button (MailCompose.xml)
    /// asks for a number of days; this module tags the outgoing email with a hidden
    /// note (ItemProps, never UserProperties - see ItemProps for the winmail.dat
    /// reason). On startup - deferred, never during it - and on a timer, tagged items
    /// in Sent Items are checked for a reply in the same conversation; unanswered ones
    /// surface in the main window's Follow-ups list, and can be sent to Todo Kanban
    /// via phase 4.
    ///
    /// Order 200: after pre-send checks (so a "Go back" leaves nothing tagged), before
    /// Send later.
    /// </summary>
    internal sealed class FollowUpOnSend : IOnSendModule
    {
        public string FeatureId => Feature.FollowUp;

        public int Order => 200;

        public void OnSend(object item, SendContext context)
        {
            // Not built. Feature.FollowUp has Built = false, so this is never called.
        }
    }
}
