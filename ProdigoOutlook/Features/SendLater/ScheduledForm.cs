using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ProdigoOutlook.Core;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// Prodigo > Scheduled: everything waiting in the Outbox to be delivered later, with
    /// Change time, Send now, Cancel and Open - so there is no need to visit the Outbox,
    /// where opening an item is what takes it out of the send queue.
    ///
    /// A meeting's row acts on the meeting, not on the waiting message: Change time and
    /// Send now resend the meeting from the calendar (which replaces the waiting
    /// invite), and Open opens the meeting in the calendar.
    /// </summary>
    internal sealed class ScheduledForm : Form
    {
        private ListView _list;
        private Button _putBack, _change, _sendNow, _cancelSend, _open;
        private List<OutboxEntry> _entries = new List<OutboxEntry>();

        public static void Open(Office.IRibbonControl control)
        {
            using (var form = new ScheduledForm())
            {
                form.ShowDialog(Ol.OwnerWindow(Ol.ContextWindow(control)));
            }
        }

        private ScheduledForm()
        {
            FormStyle.ApplyDialog(this, "Scheduled Sends");
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(880, 300);
            Build();
            Load += (s, e) => Reload();
        }

        private void Build()
        {
            ClientSize = new Size(840, 360);

            _list = new ListView
            {
                View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false,
                Left = 12, Top = 12, Width = 816, Height = 290,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            _list.Columns.Add("Delivers", 120);
            _list.Columns.Add("Status", 120);
            _list.Columns.Add("Type", 90);
            _list.Columns.Add("Subject", 280);
            _list.Columns.Add("To", 200);
            _list.SelectedIndexChanged += (s, e) => UpdateButtons();
            _list.DoubleClick += (s, e) => Act(OpenEntry);
            Controls.Add(_list);

            int top = 314;
            _putBack = AddButton("&Put back", 12, top, () => Act(PutBack));
            _change = AddButton("&Change time…", 126, top, () => Act(ChangeTime));
            _sendNow = AddButton("Send &now", 240, top, () => Act(SendNow));
            _cancelSend = AddButton("Ca&ncel send", 354, top, () => Act(CancelSend));
            _open = AddButton("&Open", 468, top, () => Act(OpenEntry));
            AddButton("&Refresh", 582, top, Reload);
            Button close = AddButton("Close", 732, top, Close);
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            CancelButton = close;

            UpdateButtons();
        }

        private Button AddButton(string text, int left, int top, Action click)
        {
            var b = new Button
            {
                Text = text, Left = left, Top = top, Width = 108, Height = 30,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            };
            b.Click += (s, e) => click();
            Controls.Add(b);
            return b;
        }

        private void Reload()
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                // Stuck items first, each on its own row - they are the ones that need
                // something done. Then one row per meeting (its latest waiting message)
                // and one per email.
                List<OutboxEntry> all = SendLaterOutbox.Waiting(includeStuck: true);
                List<OutboxEntry> queued = all.Where(e => !e.IsStuck).ToList();
                _entries = all.Where(e => e.IsStuck)
                    .Concat(queued.Where(e => !e.IsMeeting || e.MeetingId == null)
                        .Concat(queued.Where(e => e.IsMeeting && e.MeetingId != null)
                                      .GroupBy(e => e.MeetingId)
                                      .Select(g => g.OrderBy(e => e.DeliverAt).Last()))
                        .OrderBy(e => e.DeliverAt))
                    .ToList();

                _list.BeginUpdate();
                _list.Items.Clear();
                foreach (OutboxEntry e in _entries)
                {
                    var row = new ListViewItem(e.HasDeliveryTime ? TimeHelpers.Friendly(e.DeliverAt) : "-") { Tag = e };
                    row.SubItems.Add(StatusOf(e));
                    row.SubItems.Add(e.IsMeeting ? "Meeting " + e.KindLabel.ToLowerInvariant() : e.KindLabel);
                    row.SubItems.Add(e.Subject ?? string.Empty);
                    row.SubItems.Add(e.To ?? string.Empty);
                    if (e.IsStuck) row.ForeColor = Color.Firebrick;
                    _list.Items.Add(row);
                }
                _list.EndUpdate();

                int stuck = _entries.Count(e => e.IsStuck);
                if (_entries.Count == 0)
                    Text = "Scheduled Sends - nothing is waiting";
                else if (stuck > 0)
                    Text = "Scheduled Sends - " + stuck + " not queued, " + (_entries.Count - stuck) + " waiting";
                else
                    Text = "Scheduled Sends - " + _entries.Count + " waiting";
            }
            catch (Exception ex)
            {
                Log.Error("Scheduled: could not read the Outbox", ex);
                MessageBox.Show(this, "The Outbox could not be read.\n\n" + ex.Message, "Prodigo - Scheduled Sends",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                Cursor = Cursors.Default;
                UpdateButtons();
            }
        }

        private OutboxEntry Selected =>
            _list.SelectedItems.Count == 1 ? (OutboxEntry)_list.SelectedItems[0].Tag : null;

        private static string StatusOf(OutboxEntry e)
        {
            if (!e.IsStuck) return "Waiting";
            return e.HasDeliveryTime && e.DeliverAt > DateTime.Now ? "Not queued" : "Not sent (overdue)";
        }

        private void UpdateButtons()
        {
            OutboxEntry e = Selected;
            bool any = e != null;
            _cancelSend.Enabled = _open.Enabled = any;
            _putBack.Enabled = any && e.IsStuck;

            // A meeting invite delayed with Outlook's own Delay Delivery has no Prodigo
            // link to its meeting, so resending the meeting could not replace it - the
            // attendees would get two invites. A stuck meeting invite is the same: the
            // waiting-invite search does not see it, so Put back is the way to fix it.
            // Cancel and Open still work.
            bool canResend = any && (!e.IsMeeting || (e.MeetingId != null && !e.IsStuck));
            _change.Enabled = _sendNow.Enabled = canResend;
        }

        /// <summary>
        /// Puts a stuck item back in the send queue, unchanged. If its time has passed it
        /// goes straight away - after asking, for an email.
        /// </summary>
        private void PutBack(OutboxEntry entry)
        {
            bool overdue = !entry.HasDeliveryTime || entry.DeliverAt <= DateTime.Now;
            if (overdue && !Confirm("The delivery time for \"" + entry.Subject + "\" has passed. Send it now?"))
                return;

            object item = Ol.ItemById(entry.EntryId);
            if (item == null) throw new InvalidOperationException("The item is no longer in the Outbox.");

            SendLaterOutbox.BeginResubmit(entry.EntryId);
            try
            {
                if (item is Outlook.MailItem mail)
                {
                    if (overdue) SendLaterOutbox.ClearDeferred(mail);
                    mail.Send();
                }
                else if (item is Outlook.MeetingItem meeting)
                {
                    if (overdue) SendLaterOutbox.ClearDeferred(meeting);
                    meeting.Send();
                }
                Log.Info("Scheduled: put \"" + entry.Subject + "\" back in the send queue" + (overdue ? ", sending now" : ""));
            }
            finally
            {
                SendLaterOutbox.EndResubmit(entry.EntryId);
                ComScope.Release(item);
            }
        }

        /// <summary>Runs an action on the selected row, then refreshes the list.</summary>
        private void Act(Action<OutboxEntry> action)
        {
            OutboxEntry entry = Selected;
            if (entry == null) return;
            try
            {
                action(entry);
            }
            catch (Exception ex)
            {
                Log.Error("Scheduled: action failed", ex);
                MessageBox.Show(this,
                    "That did not work.\n\n" + ex.Message +
                    "\n\nIf this keeps happening for an email, open it from the Outbox and use Send Later there.",
                    "Prodigo - Scheduled Sends", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            Reload();

            // Replacing a meeting's waiting invite happens a few seconds later, so look again then.
            if (entry.IsMeeting)
                Later.Run(TimeSpan.FromSeconds(8), "Scheduled refresh", () => { if (!IsDisposed && Visible) Reload(); });
        }

        private void ChangeTime(OutboxEntry entry)
        {
            WithItem(entry,
                mail =>
                {
                    SendLaterChoice choice = SendLaterForm.Ask(this, mail.Subject, entry.DeliverAt, null, isMeeting: false);
                    if (choice == null) return;
                    SendLaterCommands.ApplyToMail(mail, choice);
                    mail.Send();
                },
                appt =>
                {
                    List<string> rooms = OutlookSettings.Current.SendLaterRoomWarning ? SendLaterOutbox.RoomNames(appt) : null;
                    DateTime? until = ItemProps.GetTime(appt, Stamp.Until);
                    DateTime shown = until.HasValue && until.Value > DateTime.Now ? until.Value : entry.DeliverAt;
                    SendLaterChoice choice = SendLaterForm.Ask(this, appt.Subject, shown, rooms, isMeeting: true);
                    if (choice == null) return;
                    SendLaterCommands.ScheduleMeeting(appt, choice);
                });
        }

        private void SendNow(OutboxEntry entry)
        {
            WithItem(entry,
                mail =>
                {
                    SendLaterCommands.ApplyToMail(mail, SendLaterChoice.Now());
                    mail.Send();
                },
                appt => SendLaterCommands.ScheduleMeeting(appt, SendLaterChoice.Now()));
        }

        private void CancelSend(OutboxEntry entry)
        {
            if (!entry.IsMeeting || entry.MeetingId == null)
            {
                if (Confirm("Delete \"" + entry.Subject + "\" from the Outbox? It will not be sent."))
                    SendLaterOutbox.DeleteByEntryId(entry.EntryId);
                return;
            }

            if (!Confirm("Cancel the meeting \"" + entry.Subject + "\"?\n\nThe invite has not been delivered, so nobody " +
                         "will hear about it. The meeting is also removed from your calendar."))
                return;

            // The waiting messages go first, the meeting second: if deleting the meeting
            // made Outlook send a cancellation, Send later treats it as replacing them.
            Outlook.AppointmentItem appt = AppointmentOf(entry);
            if (!SendLaterOutbox.DeleteAllWaiting(entry.MeetingId))
                throw new InvalidOperationException("Not every waiting invite could be removed from the Outbox.");
            if (appt != null)
            {
                try { appt.Delete(); } finally { ComScope.Release(appt); }
            }
            Log.Info("Scheduled: cancelled meeting \"" + entry.Subject + "\" before delivery");
        }

        private void OpenEntry(OutboxEntry entry)
        {
            WithItem(entry,
                mail => mail.Display(false), // OutboxGuard asks about the send queue when it closes
                appt => appt.Display(false));
        }

        private void WithItem(OutboxEntry entry, Action<Outlook.MailItem> onMail, Action<Outlook.AppointmentItem> onMeeting)
        {
            if (entry.IsMeeting)
            {
                Outlook.AppointmentItem appt = AppointmentOf(entry);
                if (appt == null) throw new InvalidOperationException("The meeting could not be found in your calendar.");
                try { onMeeting(appt); } finally { ComScope.Release(appt); }
                return;
            }

            var mail = Ol.ItemById(entry.EntryId) as Outlook.MailItem;
            if (mail == null) throw new InvalidOperationException("The email is no longer in the Outbox.");
            try { onMail(mail); } finally { ComScope.Release(mail); }
        }

        private static Outlook.AppointmentItem AppointmentOf(OutboxEntry entry)
        {
            var meeting = Ol.ItemById(entry.EntryId) as Outlook.MeetingItem;
            if (meeting == null) return null;
            try
            {
                return meeting.GetAssociatedAppointment(false);
            }
            finally
            {
                ComScope.Release(meeting);
            }
        }

        private bool Confirm(string question) =>
            MessageBox.Show(this, question, "Prodigo - Scheduled Sends", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
    }
}
