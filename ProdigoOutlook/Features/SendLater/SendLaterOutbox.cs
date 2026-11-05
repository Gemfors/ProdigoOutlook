using System;
using System.Collections.Generic;
using System.Linq;
using ProdigoOutlook.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// The names of Send later's hidden notes. On the appointment: Id, Until, Queued.
    /// On each outgoing meeting message: Id, Gen, Kind, Until. On a scheduled email: Until.
    /// </summary>
    internal static class Stamp
    {
        /// <summary>Links an appointment to every message Prodigo sends for it.</summary>
        public const string Id = "SendLater.Id";
        /// <summary>When the invite should arrive (UTC text).</summary>
        public const string Until = "SendLater.Until";
        /// <summary>"1" once the first invite has gone to the Outbox.</summary>
        public const string Queued = "SendLater.Queued";
        /// <summary>Unique per outgoing message, so one send can be told from another.</summary>
        public const string Gen = "SendLater.Gen";
        /// <summary>The MeetingMessageKind of an outgoing message.</summary>
        public const string Kind = "SendLater.Kind";
        /// <summary>Who an outgoing meeting message goes to: lower-cased addresses separated by ";".</summary>
        public const string To = "SendLater.To";
    }

    /// <summary>
    /// What Send later knows and does outside the moment of sending: in-memory notes
    /// between the form and ItemSend, and everything that reads or changes the Outbox.
    /// </summary>
    internal static class SendLaterOutbox
    {
        // ---------------------------------------------------------------
        // In-memory notes
        // ---------------------------------------------------------------

        private static readonly Dictionary<string, KeyValuePair<PendingCommand, DateTime>> Commands =
            new Dictionary<string, KeyValuePair<PendingCommand, DateTime>>();
        private static readonly Dictionary<string, CleanupBatch> Batches = new Dictionary<string, CleanupBatch>();
        private static readonly HashSet<string> Resubmitting = new HashSet<string>();

        /// <summary>
        /// A command older than this is ignored. ItemSend runs within the Send() call
        /// that follows SetCommand, so anything older belongs to a send that never
        /// happened - and must not decide the timing of some later, unrelated update.
        /// </summary>
        private static readonly TimeSpan CommandLifetime = TimeSpan.FromSeconds(30);

        /// <summary>The Send later form is about to send this meeting; tell ItemSend how.</summary>
        public static void SetCommand(string id, PendingCommand command) =>
            Commands[id] = new KeyValuePair<PendingCommand, DateTime>(command, DateTime.Now);

        public static void ClearCommand(string id)
        {
            if (id != null) Commands.Remove(id);
        }

        public static PendingCommand TakeCommand(string id)
        {
            KeyValuePair<PendingCommand, DateTime> entry;
            if (id == null || !Commands.TryGetValue(id, out entry)) return PendingCommand.None;
            Commands.Remove(id);
            return DateTime.Now - entry.Value <= CommandLifetime ? entry.Key : PendingCommand.None;
        }

        /// <summary>
        /// Marks an item as being put back in the send queue unchanged (OutboxGuard), so
        /// ItemSend leaves its timing alone.
        /// </summary>
        public static void BeginResubmit(string entryId) { if (entryId != null) Resubmitting.Add(entryId); }

        public static void EndResubmit(string entryId) { if (entryId != null) Resubmitting.Remove(entryId); }

        public static bool IsResubmitting(string entryId) => entryId != null && Resubmitting.Contains(entryId);

        // ---------------------------------------------------------------
        // Setting the delivery time
        // ---------------------------------------------------------------

        /// <summary>
        /// Sets "do not deliver before" on an email or meeting message, and checks that
        /// it took. The brief asked for exactly this: if the object model's
        /// DeferredDeliveryTime does not stick on a meeting request, write the MAPI
        /// property behind it directly. Both outcomes are logged, so the test plan can
        /// read from the log which route worked.
        /// </summary>
        public static void ApplyDeferred(object item, DateTime local)
        {
            local = TimeHelpers.TruncateToMinute(local);
            string what = item is Outlook.MeetingItem ? "meeting message" : "email";

            try
            {
                if (item is Outlook.MailItem mail) mail.DeferredDeliveryTime = local;
                else if (item is Outlook.MeetingItem meeting) meeting.DeferredDeliveryTime = local;
            }
            catch (Exception ex)
            {
                Log.Warn("DeferredDeliveryTime could not be set on " + what, ex);
            }

            DateTime readBack = ReadDeferred(item);
            if (!TimeHelpers.IsNoDate(readBack) && TimeHelpers.SameMinute(readBack, local))
            {
                Log.Info("Deferred " + what + " to " + TimeHelpers.Friendly(local) + " (DeferredDeliveryTime)");
                return;
            }

            Log.Warn("DeferredDeliveryTime did not stick on " + what + " (read back " + readBack +
                     "); falling back to PR_DEFERRED_SEND_TIME");
            if (ItemProps.SetDeferredSendTimeRaw(item, local))
            {
                DateTime again = ReadDeferred(item);
                Log.Info("Deferred " + what + " to " + TimeHelpers.Friendly(local) +
                         " (PR_DEFERRED_SEND_TIME; reads back as " + again + ")");
            }
        }

        public static void ClearDeferred(object item)
        {
            try
            {
                if (item is Outlook.MailItem mail) mail.DeferredDeliveryTime = TimeHelpers.OutlookNoDate;
                else if (item is Outlook.MeetingItem meeting) meeting.DeferredDeliveryTime = TimeHelpers.OutlookNoDate;
            }
            catch (Exception ex)
            {
                Log.Warn("Could not clear DeferredDeliveryTime", ex);
            }
        }

        public static DateTime ReadDeferred(object item)
        {
            try
            {
                if (item is Outlook.MailItem mail) return mail.DeferredDeliveryTime;
                if (item is Outlook.MeetingItem meeting) return meeting.DeferredDeliveryTime;
            }
            catch
            {
                // Fall through.
            }
            return TimeHelpers.OutlookNoDate;
        }

        // ---------------------------------------------------------------
        // Reading the Outbox
        // ---------------------------------------------------------------
        //
        // NEVER open an Outbox item just to look at it. Opening one through the object
        // model (Items[i], GetItemFromID) takes it out of the send queue exactly as
        // double-clicking it does - found the hard way on 4 November 2026, when merely
        // opening the Scheduled window left a scheduled email stuck forever. Everything
        // here reads the folder's table of contents (Folder.GetTable) instead, which
        // lists properties without opening the messages. Open an item only to act on
        // it, and then send or delete it straight away.

        private const string PrSubject = "http://schemas.microsoft.com/mapi/proptag/0x0037001F";
        private const string PrDisplayTo = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F";
        private const string PrMessageClass = "http://schemas.microsoft.com/mapi/proptag/0x001A001F";

        // Column order in the table; Row values are 1-based in this order.
        private static readonly string[] OutboxColumns =
        {
            "EntryID",                     // 1
            PrSubject,                     // 2
            PrDisplayTo,                   // 3
            PrMessageClass,                // 4
            ItemProps.DeferredDeliveryTime,// 5
            ItemProps.MessageFlags,        // 6
            ItemProps.TableSchema(Stamp.Id),    // 7
            ItemProps.TableSchema(Stamp.Gen),   // 8
            ItemProps.TableSchema(Stamp.Kind),  // 9
            ItemProps.TableSchema(Stamp.To),    // 10
        };

        /// <summary>
        /// What is in the Outbox to be delivered later - Prodigo's and anything delayed
        /// with Outlook's own Delay Delivery - read without opening anything. Every
        /// mailbox's Outbox, not only the default one, so a meeting sent from a second
        /// account is found too. Optionally only one meeting's messages.
        ///
        /// By default only items still queued and still due. With includeStuck, also
        /// items that are no longer queued (opened and closed without Send) - those sit
        /// in the Outbox forever unless someone puts them back.
        /// </summary>
        public static List<OutboxEntry> Waiting(string meetingId = null, bool includeStuck = false)
        {
            var result = new List<OutboxEntry>();
            DateTime now = DateTime.Now;

            using (var com = new ComScope())
            {
                foreach (Outlook.MAPIFolder outbox in Ol.AllFolders(Outlook.OlDefaultFolders.olFolderOutbox, com))
                {
                    try
                    {
                        foreach (OutboxEntry entry in ReadTable(com, outbox))
                        {
                            if (meetingId != null && entry.MeetingId != meetingId) continue;

                            if (entry.IsStuck)
                            {
                                if (includeStuck) result.Add(entry);
                            }
                            else if (entry.HasDeliveryTime && entry.DeliverAt > now)
                            {
                                result.Add(entry);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warn("Could not read an Outbox", ex);
                    }
                }
            }

            return result.OrderBy(e => e.HasDeliveryTime ? e.DeliverAt : DateTime.MaxValue).ToList();
        }

        public static bool HasWaiting(string meetingId) => meetingId != null && Waiting(meetingId).Count > 0;

        /// <summary>
        /// One Outbox item by EntryID, queued or stuck, read without opening it. Null when
        /// it is no longer in any Outbox (sent or deleted).
        /// </summary>
        public static OutboxEntry FindInOutbox(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return null;
            using (var com = new ComScope())
            {
                foreach (Outlook.MAPIFolder outbox in Ol.AllFolders(Outlook.OlDefaultFolders.olFolderOutbox, com))
                {
                    try
                    {
                        OutboxEntry found = ReadTable(com, outbox)
                            .FirstOrDefault(e => string.Equals(e.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
                        if (found != null) return found;
                    }
                    catch (Exception ex)
                    {
                        Log.Warn("Could not read an Outbox", ex);
                    }
                }
            }
            return null;
        }

        private static List<OutboxEntry> ReadTable(ComScope com, Outlook.MAPIFolder folder)
        {
            var entries = new List<OutboxEntry>();
            Outlook.Table table = com.Track(folder.GetTable("", Outlook.OlTableContents.olUserItems));
            Outlook.Columns columns = com.Track(table.Columns);
            columns.RemoveAll();
            foreach (string column in OutboxColumns) columns.Add(column);

            while (!table.EndOfTable)
            {
                Outlook.Row row = table.GetNextRow();
                try
                {
                    string cls = row[4] as string ?? string.Empty;
                    bool isMeeting = cls.StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase);
                    bool isMail = cls.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase);
                    if (!isMeeting && !isMail) continue;

                    var entry = new OutboxEntry
                    {
                        EntryId = row[1] as string,
                        Subject = row[2] as string,
                        To = row[3] as string,
                        IsMeeting = isMeeting,
                    };

                    if (row[5] != null)
                    {
                        DateTime local = row.UTCToLocalTime(5);
                        if (!TimeHelpers.IsNoDate(local))
                        {
                            entry.DeliverAt = local;
                            entry.HasDeliveryTime = true;
                        }
                    }

                    entry.IsStuck = !(row[6] is int flags) || (flags & ItemProps.MsgFlagSubmit) == 0;

                    if (isMeeting)
                    {
                        entry.MeetingId = row[7] as string;
                        entry.Gen = row[8] as string;
                        string to = row[10] as string;
                        entry.Recipients = to == null ? null : to.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        entry.Kind = KindFromText(row[9] as string)
                                     ?? (cls.IndexOf("Canceled", StringComparison.OrdinalIgnoreCase) >= 0
                                         ? MeetingMessageKind.CancelWholeMeeting
                                         : MeetingMessageKind.Request);
                        entry.KindLabel = entry.Kind == MeetingMessageKind.Request ? "Invite" : "Cancellation";
                    }
                    else
                    {
                        entry.KindLabel = "Email";
                    }

                    // Not queued and with no delivery time at all: an ordinary unsent draft
                    // someone parked in the Outbox. Not Send later's business.
                    if (entry.IsStuck && !entry.HasDeliveryTime) continue;

                    entries.Add(entry);
                }
                finally
                {
                    ComScope.Release(row);
                }
            }

            return entries;
        }

        /// <summary>
        /// Who a message goes to, as lower-cased addresses - the identity used to tell
        /// whether a new invite covers everyone an old one went to. Addresses rather than
        /// names, since two people can share a name.
        /// </summary>
        public static List<string> RecipientAddresses(Outlook.Recipients recipients)
        {
            var result = new List<string>();
            try
            {
                int n = recipients.Count;
                for (int i = 1; i <= n; i++)
                {
                    Outlook.Recipient r = recipients[i];
                    try
                    {
                        string address = r.Address;
                        if (string.IsNullOrWhiteSpace(address)) address = r.Name;
                        if (!string.IsNullOrWhiteSpace(address)) result.Add(address.Trim().ToLowerInvariant());
                    }
                    finally
                    {
                        ComScope.Release(r);
                    }
                }
            }
            finally
            {
                ComScope.Release(recipients);
            }
            return result;
        }

        public static MeetingMessageKind? KindFromText(string text)
        {
            MeetingMessageKind kind;
            return Enum.TryParse(text, out kind) ? kind : (MeetingMessageKind?)null;
        }

        // ---------------------------------------------------------------
        // Replacing waiting invites
        // ---------------------------------------------------------------

        /// <summary>Records a message just sent for a meeting, as part of the current batch.</summary>
        public static void NoteSent(string meetingId, string gen, MeetingMessageKind kind, bool immediate, List<string> recipients,
            string appointmentEntryId)
        {
            CleanupBatch batch;
            if (!Batches.TryGetValue(meetingId, out batch))
                Batches[meetingId] = batch = new CleanupBatch();
            if (appointmentEntryId != null) batch.AppointmentEntryId = appointmentEntryId;

            switch (kind)
            {
                case MeetingMessageKind.Request:
                    batch.RequestGen = gen;
                    batch.RequestRecipients = recipients == null ? null : WaitingMessage.RecipientSet(recipients);
                    break;
                case MeetingMessageKind.CancelWholeMeeting: batch.WholeCancelGen = gen; break;
                default: batch.RemovedCancelGens.Add(gen); break;
            }
            batch.Immediate |= immediate;
        }

        /// <summary>
        /// Removes the older messages a new send replaces, a few seconds after it - once
        /// Outlook has finished putting the new one in the Outbox. One press of Send
        /// update can send two messages (the update, and a cancellation to anyone
        /// removed), so the work waits until both have been seen.
        /// </summary>
        public static void ScheduleCleanup(string meetingId)
        {
            Later.Run(TimeSpan.FromSeconds(3), "Send later cleanup", () => RunCleanup(meetingId, 1));
        }

        private const int CleanupAttempts = 6;

        private static void RunCleanup(string meetingId, int attempt)
        {
            CleanupBatch batch;
            if (!Batches.TryGetValue(meetingId, out batch)) return; // an earlier timer already did it

            List<OutboxEntry> waiting = Waiting(meetingId).Where(e => e.Gen != null).ToList();
            bool proof = batch.ProofGen != null &&
                         (waiting.Any(e => e.Gen == batch.ProofGen) || SeenAnywhere(batch.ProofGen, batch.Immediate));

            CleanupPlan plan = SendLaterTime.PlanCleanup(
                waiting.Select(e => new WaitingMessage(e.Gen, e.Kind, e.Recipients)).ToList(), batch, proof);

            if (plan == null)
            {
                if (attempt < CleanupAttempts)
                {
                    Later.Run(TimeSpan.FromSeconds(3), "Send later cleanup", () => RunCleanup(meetingId, attempt + 1));
                    return;
                }
                Batches.Remove(meetingId);
                Log.Warn("Send later: the new message for meeting " + meetingId + " never showed up in the Outbox; " +
                         "older waiting messages were left alone, so attendees get the invite first and then the change.");
                return;
            }

            Batches.Remove(meetingId);

            var byGen = waiting.GroupBy(e => e.Gen).ToDictionary(g => g.Key, g => g.First().EntryId);
            int deleted = 0;
            foreach (string gen in plan.First)
            {
                if (!DeleteByEntryId(byGen[gen]))
                {
                    Log.Warn("Send later: could not remove a waiting message; stopping, so nothing is left half-replaced.");
                    return;
                }
                deleted++;
            }
            foreach (string gen in plan.Then)
            {
                if (DeleteByEntryId(byGen[gen])) deleted++;
            }

            Log.Info("Send later: replaced " + deleted + " waiting message(s) for meeting " + meetingId);

            var waitingMessages = waiting.Select(e => new WaitingMessage(e.Gen, e.Kind, e.Recipients)).ToList();
            if (SendLaterTime.NeedsUpdateForRemoved(waitingMessages, batch, plan))
                SendUpdateForRemoved(meetingId, batch);
        }

        /// <summary>
        /// Sends the meeting again to the attendees still on it, so that the waiting
        /// invite (which still lists the removed people) and the cancellation to them can
        /// both be replaced. The cancellation is carried into the new batch, so the
        /// cleanup after this send counts the removed people as accounted for. If
        /// anything fails, nothing is deleted: everyone gets the original invite, and
        /// the removed people the cancellation after it - the same as without this.
        /// </summary>
        private static void SendUpdateForRemoved(string meetingId, CleanupBatch earlier)
        {
            if (string.IsNullOrEmpty(earlier.AppointmentEntryId)) return;

            var carried = new CleanupBatch { AppointmentEntryId = earlier.AppointmentEntryId };
            carried.RemovedCancelGens.AddRange(earlier.RemovedCancelGens);
            Batches[meetingId] = carried;

            Later.Run(TimeSpan.FromSeconds(1), "Send later update for removed attendees", () =>
            {
                var appt = Ol.ItemById(carried.AppointmentEntryId) as Outlook.AppointmentItem;
                try
                {
                    if (appt == null) throw new InvalidOperationException("appointment not found");
                    Log.Info("Send later: someone was removed while the invite waits; sending the update to the remaining attendees");
                    appt.Send();
                }
                catch (Exception ex)
                {
                    Log.Warn("Send later: could not send the update for removed attendees; the original invite and the cancellation stay", ex);
                    if (Batches.TryGetValue(meetingId, out CleanupBatch b) && b == carried) Batches.Remove(meetingId);
                }
                finally
                {
                    ComScope.Release(appt);
                }
            });

            // If the update never reaches ItemSend, don't leave the carried batch to be
            // picked up by some later, unrelated send.
            Later.Run(TimeSpan.FromSeconds(60), "Send later forget carried batch", () =>
            {
                if (Batches.TryGetValue(meetingId, out CleanupBatch b) && b == carried && b.RequestGen == null)
                    Batches.Remove(meetingId);
            });
        }

        /// <summary>Looks for a message by its Gen in the whole Outbox, or in recent Sent Items.</summary>
        private static bool SeenAnywhere(string gen, bool includeSentItems)
        {
            using (var com = new ComScope())
            {
                try
                {
                    foreach (Outlook.MAPIFolder outbox in Ol.AllFolders(Outlook.OlDefaultFolders.olFolderOutbox, com))
                        if (FolderHasGen(com, outbox, gen)) return true;

                    if (!includeSentItems) return false;
                    foreach (Outlook.MAPIFolder sent in Ol.AllFolders(Outlook.OlDefaultFolders.olFolderSentMail, com))
                        if (FolderHasGen(com, sent, gen)) return true;
                    return false;
                }
                catch (Exception ex)
                {
                    Log.Warn("Send later: could not look for the new message", ex);
                    return false;
                }
            }
        }

        /// <summary>
        /// True when the folder holds a message with this Gen. A filtered table, so
        /// nothing is opened (see "Reading the Outbox") and a large Sent Items folder
        /// is searched by Outlook rather than item by item.
        /// </summary>
        private static bool FolderHasGen(ComScope com, Outlook.MAPIFolder folder, string gen)
        {
            if (string.IsNullOrEmpty(gen) || gen.IndexOf('\'') >= 0) return false;
            string filter = "@SQL=\"" + ItemProps.TableSchema(Stamp.Gen) + "\" = '" + gen + "'";
            Outlook.Table table = com.Track(folder.GetTable(filter, Outlook.OlTableContents.olUserItems));
            return !table.EndOfTable;
        }

        public static bool DeleteByEntryId(string entryId)
        {
            object item = Ol.ItemById(entryId);
            if (item == null) return true; // already gone: delivered, or deleted by hand
            try
            {
                if (item is Outlook.MailItem mail) mail.Delete();
                else if (item is Outlook.MeetingItem meeting) meeting.Delete();
                return true;
            }
            catch (Exception ex)
            {
                Log.Warn("Could not delete Outbox item", ex);
                return false;
            }
            finally
            {
                ComScope.Release(item);
            }
        }

        /// <summary>Deletes every waiting message for one meeting. True when all are gone.</summary>
        public static bool DeleteAllWaiting(string meetingId)
        {
            bool ok = true;
            foreach (OutboxEntry e in Waiting(meetingId))
                ok &= DeleteByEntryId(e.EntryId);
            return ok;
        }

        // ---------------------------------------------------------------
        // Appointment helpers
        // ---------------------------------------------------------------

        /// <summary>
        /// A second or two after the first invite went out, writes "queued" on the
        /// appointment. Not done inside ItemSend: Outlook saves the appointment itself
        /// as part of sending, and a save of ours in the middle would race with it.
        /// </summary>
        public static void MarkQueuedLater(string appointmentEntryId)
        {
            UpdateAppointmentLater(appointmentEntryId, "mark queued", appt => ItemProps.Set(appt, Stamp.Queued, "1"));
        }

        public static void ClearStampLater(string appointmentEntryId)
        {
            UpdateAppointmentLater(appointmentEntryId, "clear stamp", appt =>
            {
                ItemProps.Set(appt, Stamp.Until, null);
                ItemProps.Set(appt, Stamp.Queued, null);
            });
        }

        private static void UpdateAppointmentLater(string entryId, string what, Action<Outlook.AppointmentItem> change, int attempt = 1)
        {
            if (string.IsNullOrEmpty(entryId))
            {
                Log.Warn("Send later: cannot " + what + " - the appointment has no EntryID yet");
                return;
            }

            Later.Run(TimeSpan.FromSeconds(2 * attempt), "Send later " + what, () =>
            {
                var appt = Ol.ItemById(entryId) as Outlook.AppointmentItem;
                if (appt == null) return;
                try
                {
                    change(appt);
                    appt.Save();
                }
                catch (Exception ex)
                {
                    if (attempt < 3)
                        UpdateAppointmentLater(entryId, what, change, attempt + 1);
                    else
                        Log.Warn("Send later: could not " + what + " on the appointment", ex);
                }
                finally
                {
                    ComScope.Release(appt);
                }
            });
        }

        /// <summary>Names of meeting rooms among the attendees, for the warning in the form.</summary>
        public static List<string> RoomNames(Outlook.AppointmentItem appt)
        {
            var rooms = new List<string>();
            Outlook.Recipients recipients = null;
            try
            {
                recipients = appt.Recipients;
                int n = recipients.Count;
                for (int i = 1; i <= n; i++)
                {
                    Outlook.Recipient r = recipients[i];
                    Outlook.AddressEntry entry = null;
                    try
                    {
                        bool isRoom = r.Type == (int)Outlook.OlMeetingRecipientType.olResource;
                        if (!isRoom)
                        {
                            entry = r.AddressEntry;
                            object displayType = entry == null ? null : ItemProps.GetRaw(entry, ItemProps.DisplayTypeEx);
                            isRoom = displayType is int && (((int)displayType) & 0xFF) == 7; // DT_ROOM
                        }
                        if (isRoom) rooms.Add(r.Name);
                    }
                    catch
                    {
                        // Unresolved recipient; not a room we can recognise.
                    }
                    finally
                    {
                        ComScope.Release(entry);
                        ComScope.Release(r);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Could not check for meeting rooms", ex);
            }
            finally
            {
                ComScope.Release(recipients);
            }
            return rooms;
        }
    }

    internal sealed class OutboxEntry
    {
        public string EntryId { get; set; }
        public DateTime DeliverAt { get; set; }
        public bool HasDeliveryTime { get; set; }
        /// <summary>In the Outbox but no longer queued: it will never be sent unless put back.</summary>
        public bool IsStuck { get; set; }
        public string Subject { get; set; }
        public string To { get; set; }
        public bool IsMeeting { get; set; }
        public string MeetingId { get; set; }
        public string Gen { get; set; }
        /// <summary>For Prodigo-tagged meeting messages: lower-cased addresses. Null when unknown.</summary>
        public List<string> Recipients { get; set; }
        public MeetingMessageKind Kind { get; set; }
        public string KindLabel { get; set; }
    }
}
