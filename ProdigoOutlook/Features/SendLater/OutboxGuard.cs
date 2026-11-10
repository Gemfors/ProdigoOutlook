using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ProdigoOutlook.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// Protects scheduled items from Outlook's Outbox trap.
    ///
    /// Opening an item in the Outbox takes it out of the send queue. If the window is
    /// then closed without pressing Send, the item sits in the Outbox and is never
    /// delivered - silently. This watches every window opened on a scheduled Outbox
    /// item and, when it closes unsent:
    /// - an email: asks whether to keep it scheduled (Yes puts it back in the queue);
    /// - a meeting invite: puts it back without asking, since an invite left unsent is
    ///   never what anyone wants, and the person is pointed to the calendar to edit it.
    ///
    /// Covers items delayed with Outlook's own Delay Delivery too, not only Prodigo's.
    /// </summary>
    internal static class OutboxGuard
    {
        // Held so the event subscriptions live as long as the add-in. An event source
        // that only a local variable refers to is garbage-collected, and its events
        // quietly stop arriving.
        private static Outlook.Inspectors _inspectors;
        private static readonly List<Watch> Watches = new List<Watch>();

        public static void Start(Outlook.Application app)
        {
            _inspectors = app.Inspectors;
            _inspectors.NewInspector += OnNewInspector;
        }

        private static void OnNewInspector(Outlook.Inspector inspector)
        {
            try
            {
                if (!Feature.IsOn(Feature.SendLater)) return;

                // The window's own item: not released here, the window is about to use it.
                object item = inspector.CurrentItem;
                {
                    if (!(item is Outlook.MailItem) && !(item is Outlook.MeetingItem)) return;

                    DateTime deliver = SendLaterOutbox.ReadDeferred(item);
                    if (TimeHelpers.IsNoDate(deliver) || deliver <= DateTime.Now) return;
                    if (!Ol.IsInOutbox(item)) return;

                    var watch = new Watch(inspector, SendLaterOnSend.SafeEntryId(item), item is Outlook.MeetingItem,
                        SubjectOf(item), deliver);
                    Watches.Add(watch);

                    if (watch.IsMeeting)
                        Later.Run(TimeSpan.FromMilliseconds(600), "Outbox meeting hint", () => watch.ShowHint());
                }
            }
            catch (Exception ex)
            {
                Log.Warn("OutboxGuard: could not check a newly opened window", ex);
            }
        }

        private static string SubjectOf(object item)
        {
            try
            {
                if (item is Outlook.MailItem mail) return mail.Subject;
                if (item is Outlook.MeetingItem meeting) return meeting.Subject;
            }
            catch
            {
                // Fall through.
            }
            return string.Empty;
        }

        private sealed class Watch
        {
            private readonly Outlook.Inspector _inspector;
            private readonly string _entryId;
            private readonly string _subject;
            private readonly DateTime _deliver;
            private bool _closed;

            public Watch(Outlook.Inspector inspector, string entryId, bool isMeeting, string subject, DateTime deliver)
            {
                _inspector = inspector;
                _entryId = entryId;
                IsMeeting = isMeeting;
                _subject = subject;
                _deliver = deliver;
                ((Outlook.InspectorEvents_10_Event)_inspector).Close += OnClose;
            }

            public bool IsMeeting { get; }

            public void ShowHint()
            {
                if (_closed) return;
                try
                {
                    if (_inspector.CurrentItem is Outlook.MeetingItem meeting)
                    {
                        var answer = MessageBox.Show(Ol.OwnerWindow(_inspector),
                            "This is the invite for \"" + _subject + "\", waiting to be delivered " +
                            TimeHelpers.Friendly(_deliver) + ".\n\nTo change the meeting, edit it in your calendar and " +
                            "press Send Update - Prodigo keeps the original time and replaces this invite.\n\n" +
                            "Open the meeting in your calendar now?",
                            "Prodigo - Send Later", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (answer != DialogResult.Yes) return;

                        Outlook.AppointmentItem appt = meeting.GetAssociatedAppointment(false);
                        if (appt != null)
                        {
                            try { appt.Display(false); } finally { ComScope.Release(appt); }
                        }
                        _inspector.Close(Outlook.OlInspectorClose.olDiscard);
                    }
                }
                catch (Exception ex)
                {
                    Log.Warn("OutboxGuard: meeting hint failed", ex);
                }
            }

            private void OnClose()
            {
                _closed = true;
                try { ((Outlook.InspectorEvents_10_Event)_inspector).Close -= OnClose; } catch { }
                Watches.Remove(this);

                // After the window has really gone and Outlook has let go of the item.
                Later.Run(TimeSpan.FromMilliseconds(1500), "Outbox guard check", Check);
            }

            private void Check()
            {
                // Looked up in the Outbox's table first, without opening the item: if
                // the user pressed Send in the window, opening it now would take it
                // straight back out of the queue. Only a stuck item is opened, and only
                // to send it.
                OutboxEntry entry = SendLaterOutbox.FindInOutbox(_entryId);
                if (entry == null || !entry.IsStuck) return; // sent, deleted, or queued as it should be

                DateTime deliver = entry.DeliverAt;
                bool stillAhead = entry.HasDeliveryTime && deliver > DateTime.Now;

                if (!entry.IsMeeting)
                {
                    string question = stillAhead
                        ? "\"" + _subject + "\" was opened from the Outbox, which takes it out of the send queue. " +
                          "It will not be sent unless it is put back.\n\nKeep it scheduled for " + TimeHelpers.Friendly(deliver) + "?"
                        : "\"" + _subject + "\" was opened from the Outbox and is no longer queued, and its delivery time has passed.\n\nSend it now?";

                    var answer = MessageBox.Show(Ol.OwnerWindow(), question, "Prodigo - Send Later",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if (answer != DialogResult.Yes)
                    {
                        Log.Info("OutboxGuard: user left " + Log.Item(_entryId) + " unsent in the Outbox");
                        return;
                    }
                }

                object item = Ol.ItemById(_entryId);
                if (item == null) return;

                try
                {

                    if (item is Outlook.MeetingItem meeting)
                    {
                        Resubmit(meeting, () => meeting.Send());
                        Log.Info("OutboxGuard: put the invite " + Log.Item(_entryId) + " back in the send queue");
                        return;
                    }

                    if (item is Outlook.MailItem mail)
                    {
                        Resubmit(mail, () => mail.Send());
                        Log.Info("OutboxGuard: put the email " + Log.Item(_entryId) + " back in the send queue");
                    }
                }
                catch (Exception ex)
                {
                    Log.Warn("OutboxGuard: could not put an item back in the send queue", ex);
                }
                finally
                {
                    ComScope.Release(item);
                }
            }

            private void Resubmit(object item, Action send)
            {
                SendLaterOutbox.BeginResubmit(_entryId);
                try
                {
                    send();
                }
                finally
                {
                    SendLaterOutbox.EndResubmit(_entryId);
                }
            }
        }
    }
}
