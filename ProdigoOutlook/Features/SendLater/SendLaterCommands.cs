using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ProdigoOutlook.Core;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// What the Send Later button does, in an email window and in a meeting window.
    ///
    /// The button only writes the stamp and presses Send. Deciding the actual delivery
    /// time happens in SendLaterOnSend, during ItemSend - the one place every send
    /// passes through, whether it came from this button, Outlook's own Send or Send
    /// Update button, or a script.
    /// </summary>
    internal static class SendLaterCommands
    {
        public static void SendLater(Office.IRibbonControl control)
        {
            Outlook.Inspector inspector = Ol.ContextInspector(control);
            object item = inspector?.CurrentItem;
            if (item == null)
            {
                Ol.Info("Open an email or a meeting first.", "Send Later");
                return;
            }

            // The window's own item: not released here, the inspector is still using it.
            IWin32Window owner = Ol.OwnerWindow(inspector);
            if (item is Outlook.MailItem mail) ForMail(mail, owner);
            else if (item is Outlook.AppointmentItem appt) ForMeeting(appt, owner);
            else if (item is Outlook.MeetingItem meeting) EditInCalendarInstead(meeting, owner);
            else Ol.Info("Send Later works on emails and meetings.", "Send Later");
        }

        // ---------------------------------------------------------------
        // Email
        // ---------------------------------------------------------------

        public static void ForMail(Outlook.MailItem mail, IWin32Window owner)
        {
            if (mail.Sent)
            {
                Ol.Info("This email has already been sent.", "Send Later");
                return;
            }
            if (!HasRecipients(mail.Recipients))
            {
                Ol.Info("Add a recipient first.", "Send Later");
                return;
            }

            SendLaterChoice choice = SendLaterForm.Ask(owner, mail.Subject, CurrentlyScheduled(mail), null, isMeeting: false);
            if (choice == null) return;

            ApplyToMail(mail, choice);
            mail.Send();
        }

        /// <summary>Stamps (or unstamps) an email and sets its delivery time. The caller sends it.</summary>
        public static void ApplyToMail(Outlook.MailItem mail, SendLaterChoice choice)
        {
            if (choice.SendNow)
            {
                ItemProps.Set(mail, Stamp.Until, null);
                SendLaterOutbox.ClearDeferred(mail);
                Log.Info("Send later: email \"" + mail.Subject + "\" sent now");
            }
            else
            {
                ItemProps.SetTime(mail, Stamp.Until, choice.When);
                SendLaterOutbox.ApplyDeferred(mail, choice.When);
            }
        }

        /// <summary>When this email is due to go, if it is already scheduled.</summary>
        private static DateTime? CurrentlyScheduled(Outlook.MailItem mail)
        {
            DateTime now = DateTime.Now;
            DateTime deferred = mail.DeferredDeliveryTime;
            if (!TimeHelpers.IsNoDate(deferred) && deferred > now) return deferred;
            DateTime? stamp = ItemProps.GetTime(mail, Stamp.Until);
            return stamp.HasValue && stamp.Value > now ? stamp : null;
        }

        // ---------------------------------------------------------------
        // Meeting
        // ---------------------------------------------------------------

        public static void ForMeeting(Outlook.AppointmentItem appt, IWin32Window owner)
        {
            Outlook.OlMeetingStatus status = appt.MeetingStatus;
            if (status == Outlook.OlMeetingStatus.olMeetingReceived || status == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
            {
                Ol.Info("Only the organizer can schedule the invite for this meeting.", "Send Later");
                return;
            }
            if (status == Outlook.OlMeetingStatus.olNonMeeting || !HasRecipients(appt.Recipients))
            {
                Ol.Info("This appointment has no attendees yet. Invite attendees first, then use Send Later.", "Send Later");
                return;
            }
            if (status == Outlook.OlMeetingStatus.olMeetingCanceled)
            {
                Ol.Info("This meeting has been cancelled.", "Send Later");
                return;
            }

            string id = ItemProps.Get(appt, Stamp.Id);
            DateTime? current = null;
            if (id != null)
            {
                DateTime? until = ItemProps.GetTime(appt, Stamp.Until);
                bool queued = ItemProps.Get(appt, Stamp.Queued) == "1" || SendLaterOutbox.HasWaiting(id);
                if (until.HasValue && until.Value > DateTime.Now && queued) current = until;
            }

            List<string> rooms = OutlookSettings.Current.SendLaterRoomWarning
                ? SendLaterOutbox.RoomNames(appt)
                : null;

            SendLaterChoice choice = SendLaterForm.Ask(owner, appt.Subject, current, rooms, isMeeting: true);
            if (choice == null) return;

            ScheduleMeeting(appt, choice);
        }

        /// <summary>
        /// Stamps the appointment and sends it. Also used by the Scheduled window for
        /// Change time and Send now, which is why it does not show a form itself.
        /// </summary>
        public static void ScheduleMeeting(Outlook.AppointmentItem appt, SendLaterChoice choice)
        {
            string id = ItemProps.Get(appt, Stamp.Id);
            if (id == null)
            {
                id = Guid.NewGuid().ToString("N");
                ItemProps.Set(appt, Stamp.Id, id);
            }

            DateTime? previousUntil = ItemProps.GetTime(appt, Stamp.Until);
            if (choice.SendNow)
            {
                SendLaterOutbox.SetCommand(id, PendingCommand.SendNow);
            }
            else
            {
                ItemProps.SetTime(appt, Stamp.Until, choice.When);
                SendLaterOutbox.SetCommand(id, PendingCommand.Exact);
            }

            try
            {
                // Saved first so the stamp is on the stored appointment, not only on the
                // copy in this window, by the time ItemSend reads it back.
                appt.Save();
                appt.Send();
            }
            catch
            {
                // Send refused (an unresolved name, say): put things back as they were,
                // so neither the new time nor the command lingers to affect a later send.
                SendLaterOutbox.ClearCommand(id);
                try
                {
                    ItemProps.SetTime(appt, Stamp.Until, previousUntil);
                    appt.Save();
                }
                catch (Exception ex)
                {
                    Log.Warn("Send later: could not restore the previous time after a failed send", ex);
                }
                throw;
            }
        }

        /// <summary>
        /// A meeting message opened from the Outbox: the copy, not the meeting. Changing
        /// it would not change the meeting, so point to the calendar instead.
        /// </summary>
        public static void EditInCalendarInstead(Outlook.MeetingItem meeting, IWin32Window owner)
        {
            var answer = MessageBox.Show(owner,
                "This is the invite waiting to be sent, not the meeting itself.\n\n" +
                "To change the meeting or its time, open it in your calendar and use Send Update or Send Later there - " +
                "Prodigo keeps the original delivery time and replaces this invite.\n\nOpen the meeting in your calendar now?",
                "Prodigo - Send Later", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (answer != DialogResult.Yes) return;

            Outlook.AppointmentItem appt = meeting.GetAssociatedAppointment(false);
            if (appt == null)
            {
                Ol.Info("The meeting could not be found in your calendar.", "Send Later");
                return;
            }
            try
            {
                appt.Display(false);
            }
            finally
            {
                ComScope.Release(appt);
            }
        }

        private static bool HasRecipients(Outlook.Recipients recipients)
        {
            try
            {
                return recipients.Count > 0;
            }
            finally
            {
                ComScope.Release(recipients);
            }
        }
    }
}
