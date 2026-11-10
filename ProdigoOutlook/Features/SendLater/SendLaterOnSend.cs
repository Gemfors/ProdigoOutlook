using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ProdigoOutlook.Core;
using ProdigoOutlook.OnSend;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// Send later's part of ItemSend. Runs last of the on-send modules.
    ///
    /// For an email: if Prodigo scheduled it and the delivery time has gone missing
    /// (it was opened from the Outbox and sent again), put the time back.
    ///
    /// For a meeting message: read the stamp on the appointment it belongs to, decide
    /// when this message should arrive (SendLaterTime.Decide), set that, tag the
    /// message so it can be found in the Outbox later, and arrange for older waiting
    /// messages to be replaced once this one is safely queued.
    ///
    /// Never stops a send. Anything that goes wrong leaves the item to go as Outlook
    /// would have sent it.
    /// </summary>
    internal sealed class SendLaterOnSend : IOnSendModule
    {
        public string FeatureId => Feature.SendLater;

        public int Order => 300;

        public void OnSend(object item, SendContext context)
        {
            if (item is Outlook.MailItem mail)
                OnMail(mail);
            else if (item is Outlook.MeetingItem meeting)
                OnMeeting(meeting);
        }

        private static void OnMail(Outlook.MailItem mail)
        {
            if (SendLaterOutbox.IsResubmitting(SafeEntryId(mail))) return;

            DateTime now = DateTime.Now;
            DateTime? until = ItemProps.GetTime(mail, Stamp.Until);
            if (!until.HasValue || until.Value <= now) return;

            DateTime current = mail.DeferredDeliveryTime;
            bool hasTime = !TimeHelpers.IsNoDate(current) && current > now;

            if (hasTime && TimeHelpers.SameMinute(current, until.Value)) return;

            if (hasTime)
            {
                // Changed with Outlook's own Delay Delivery: the user's choice wins.
                ItemProps.SetTime(mail, Stamp.Until, current);
                Log.Info("Send later: email's delivery time was changed in Outlook to " + TimeHelpers.Friendly(current));
                return;
            }

            // The time is gone. Either Outlook dropped it when the email was sent again
            // from the Outbox, or the user removed it on purpose - Prodigo cannot tell
            // which, so it asks. Neither answer stops the send.
            var answer = MessageBox.Show(Ol.OwnerWindow(),
                "\"" + mail.Subject + "\" was scheduled for " + TimeHelpers.Friendly(until.Value) +
                ", but no longer has a delivery time.\n\nKeep it scheduled for " + TimeHelpers.Friendly(until.Value) +
                "?\n\nYes keeps the schedule. No sends it now.",
                "Prodigo - Send Later", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

            if (answer == DialogResult.Yes)
            {
                Log.Info("Send later: email had lost its delivery time; restored " + TimeHelpers.Friendly(until.Value));
                SendLaterOutbox.ApplyDeferred(mail, until.Value);
            }
            else
            {
                ItemProps.Set(mail, Stamp.Until, null);
                Log.Info("Send later: email had lost its delivery time; user chose to send it now");
            }
        }

        private static void OnMeeting(Outlook.MeetingItem meeting)
        {
            if (SendLaterOutbox.IsResubmitting(SafeEntryId(meeting))) return;

            // Not released at the end: while Prodigo's own button is sending, this is very
            // likely the same object the button still holds, and releasing it here would
            // pull it out from under that code. It is let go when it goes out of scope.
            Outlook.AppointmentItem appt = meeting.GetAssociatedAppointment(false);
            if (appt == null) return;

            {
                string id = ItemProps.Get(appt, Stamp.Id);
                if (id == null) return; // never scheduled with Prodigo

                DateTime now = DateTime.Now;
                DateTime? until = ItemProps.GetTime(appt, Stamp.Until);
                PendingCommand command = SendLaterOutbox.TakeCommand(id);
                MeetingMessageKind kind = KindOf(meeting, appt);

                List<OutboxEntry> waiting = SendLaterOutbox.Waiting(id);
                bool queued = ItemProps.Get(appt, Stamp.Queued) == "1" || waiting.Count > 0;
                DateTime? latestWaiting = waiting.Count > 0 ? waiting.Max(e => e.DeliverAt) : (DateTime?)null;

                MeetingDecision decision = SendLaterTime.Decide(kind, until, queued, command, now, latestWaiting);
                Log.Info("Send later: " + Log.Meeting(id) + ", " + kind + ", " + decision.Reason +
                         (decision.Defer ? ", deliver " + TimeHelpers.Friendly(decision.DeliverAt) : ""));

                if (!decision.Defer && !decision.CleanUp && !decision.ClearStamp) return;

                string gen = Guid.NewGuid().ToString("N");
                List<string> recipients = SendLaterOutbox.RecipientAddresses(meeting.Recipients);
                ItemProps.Set(meeting, Stamp.Id, id);
                ItemProps.Set(meeting, Stamp.Gen, gen);
                ItemProps.Set(meeting, Stamp.Kind, kind.ToString());
                ItemProps.Set(meeting, Stamp.To, string.Join(";", recipients));

                if (decision.Defer)
                {
                    ItemProps.SetTime(meeting, Stamp.Until, decision.DeliverAt);
                    SendLaterOutbox.ApplyDeferred(meeting, decision.DeliverAt);
                }

                string apptEntryId = SafeEntryId(appt);
                if (decision.MarkQueued) SendLaterOutbox.MarkQueuedLater(apptEntryId);
                if (decision.ClearStamp) SendLaterOutbox.ClearStampLater(apptEntryId);
                if (decision.CleanUp)
                {
                    SendLaterOutbox.NoteSent(id, gen, kind, immediate: !decision.Defer, recipients: recipients);
                    SendLaterOutbox.ScheduleCleanup(id);
                }
            }
        }

        /// <summary>
        /// An invite or update, a cancellation of the whole meeting, or the cancellation
        /// Outlook sends to people removed from it (the meeting itself still stands).
        /// </summary>
        private static MeetingMessageKind KindOf(Outlook.MeetingItem meeting, Outlook.AppointmentItem appt)
        {
            string cls = meeting.MessageClass ?? string.Empty;
            if (cls.IndexOf("Canceled", StringComparison.OrdinalIgnoreCase) < 0)
                return MeetingMessageKind.Request;

            // MeetingStatus alone is not enough: on 5 November 2026 Outlook's own Cancel
            // Meeting arrived here with the status not yet set to cancelled. Who the
            // cancellation goes to settles it - see SendLaterTime.CancellationKind.
            bool statusSaysCanceled = false;
            try { statusSaysCanceled = appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled; } catch { }
            return SendLaterTime.CancellationKind(statusSaysCanceled,
                SendLaterOutbox.RecipientAddresses(meeting.Recipients),
                SendLaterOutbox.RecipientAddresses(appt.Recipients));
        }

        internal static string SafeEntryId(object item)
        {
            try
            {
                if (item is Outlook.MailItem mail) return mail.EntryID;
                if (item is Outlook.MeetingItem meeting) return meeting.EntryID;
                if (item is Outlook.AppointmentItem appt) return appt.EntryID;
            }
            catch
            {
                // Not saved yet.
            }
            return null;
        }
    }
}
