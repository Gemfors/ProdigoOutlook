using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ProdigoOutlook.Core;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// Everything about Send later that can be decided without Outlook: what the
    /// presets mean, whether a time is acceptable, when a meeting message should be
    /// delivered, and which waiting invites to remove. No Outlook types, so
    /// test\LogicTests compiles this file directly - the lesson from ProdigoExcel:
    /// put the logic that can be wrong where it can be tested.
    /// </summary>
    internal static class SendLaterTime
    {
        /// <summary>A delivery time must be at least this far ahead when it is chosen.</summary>
        public static readonly TimeSpan MinimumLead = TimeSpan.FromMinutes(1);

        /// <summary>How much later than the invite an update or cancellation is delivered, if it cannot replace it.</summary>
        public static readonly TimeSpan UpdateOffset = TimeSpan.FromMinutes(1);

        // ---------------------------------------------------------------
        // Presets
        // ---------------------------------------------------------------

        private static readonly Regex Relative = new Regex(@"^\+\s*(?:(\d+)\s*h)?\s*(?:(\d+)\s*m(?:in)?)?$", RegexOptions.IgnoreCase);
        private static readonly Regex AtTime = new Regex(@"^([\p{L}]+)\s+(\d{1,2})[:.](\d{2})$", RegexOptions.IgnoreCase);

        private static readonly Dictionary<string, DayOfWeek> Weekdays = new Dictionary<string, DayOfWeek>(StringComparer.OrdinalIgnoreCase)
        {
            { "monday", DayOfWeek.Monday }, { "mon", DayOfWeek.Monday }, { "måndag", DayOfWeek.Monday }, { "mån", DayOfWeek.Monday },
            { "tuesday", DayOfWeek.Tuesday }, { "tue", DayOfWeek.Tuesday }, { "tisdag", DayOfWeek.Tuesday }, { "tis", DayOfWeek.Tuesday },
            { "wednesday", DayOfWeek.Wednesday }, { "wed", DayOfWeek.Wednesday }, { "onsdag", DayOfWeek.Wednesday }, { "ons", DayOfWeek.Wednesday },
            { "thursday", DayOfWeek.Thursday }, { "thu", DayOfWeek.Thursday }, { "torsdag", DayOfWeek.Thursday }, { "tors", DayOfWeek.Thursday },
            { "friday", DayOfWeek.Friday }, { "fri", DayOfWeek.Friday }, { "fredag", DayOfWeek.Friday }, { "fre", DayOfWeek.Friday },
            { "saturday", DayOfWeek.Saturday }, { "sat", DayOfWeek.Saturday }, { "lördag", DayOfWeek.Saturday }, { "lör", DayOfWeek.Saturday },
            { "sunday", DayOfWeek.Sunday }, { "sun", DayOfWeek.Sunday }, { "söndag", DayOfWeek.Sunday }, { "sön", DayOfWeek.Sunday },
        };

        /// <summary>
        /// Reads one preset line from the settings. Returns null for anything it does
        /// not understand, so a typo costs that one button rather than the form.
        /// <list type="bullet">
        /// <item>"+1h", "+30m", "+1h30m" - that long from now</item>
        /// <item>"today 13:00", "tomorrow 08:00" (also "idag", "imorgon")</item>
        /// <item>"nextworkday 08:00" - the next Monday to Friday after today</item>
        /// <item>"monday 08:00", "måndag 08:00", "fri 16:00" - the next such day, today included if the time is still ahead</item>
        /// </list>
        /// </summary>
        public static SendLaterPreset ParsePreset(string spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return null;
            string s = spec.Trim();

            Match rel = Relative.Match(s);
            if (rel.Success && (rel.Groups[1].Success || rel.Groups[2].Success))
            {
                int h = rel.Groups[1].Success ? int.Parse(rel.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
                int m = rel.Groups[2].Success ? int.Parse(rel.Groups[2].Value, CultureInfo.InvariantCulture) : 0;
                var span = new TimeSpan(h, m, 0);
                if (span <= TimeSpan.Zero) return null;
                return new SendLaterPreset(s, RelativeLabel(h, m), now => TimeHelpers.TruncateToMinute(now + span));
            }

            Match at = AtTime.Match(s);
            if (!at.Success) return null;

            int hour = int.Parse(at.Groups[2].Value, CultureInfo.InvariantCulture);
            int minute = int.Parse(at.Groups[3].Value, CultureInfo.InvariantCulture);
            if (hour > 23 || minute > 59) return null;
            var time = new TimeSpan(hour, minute, 0);
            string hhmm = hour.ToString("00") + ":" + minute.ToString("00");
            string word = at.Groups[1].Value.ToLowerInvariant();

            switch (word)
            {
                case "today":
                case "idag":
                    return new SendLaterPreset(s, "Today " + hhmm, now => now.Date + time);

                case "tomorrow":
                case "imorgon":
                    return new SendLaterPreset(s, "Tomorrow " + hhmm, now => now.Date.AddDays(1) + time);

                case "nextworkday":
                case "workday":
                    return new SendLaterPreset(s, "Next workday " + hhmm, now =>
                    {
                        DateTime d = now.Date.AddDays(1);
                        while (TimeHelpers.IsWeekend(d)) d = d.AddDays(1);
                        return d + time;
                    });
            }

            DayOfWeek day;
            if (Weekdays.TryGetValue(word, out day))
            {
                string name = TimeHelpers.UiCulture.DateTimeFormat.GetDayName(day);
                return new SendLaterPreset(s, name + " " + hhmm, now =>
                {
                    DateTime d = now.Date;
                    while (d.DayOfWeek != day || d + time < now + MinimumLead) d = d.AddDays(1);
                    return d + time;
                });
            }

            return null;
        }

        private static string RelativeLabel(int h, int m)
        {
            if (m == 0) return h == 1 ? "In 1 hour" : "In " + h + " hours";
            if (h == 0) return "In " + m + " minutes";
            return "In " + h + " h " + m + " min";
        }

        /// <summary>
        /// Null when the time is fine, otherwise the sentence to show under the picker.
        /// </summary>
        public static string Validate(DateTime chosen, DateTime now)
        {
            if (chosen <= now) return "That time has already passed.";
            if (chosen < now + MinimumLead) return "Pick a time at least a minute from now.";
            return null;
        }

        // ---------------------------------------------------------------
        // Meetings: when should this message be delivered?
        // ---------------------------------------------------------------

        /// <summary>
        /// Decides what happens to one meeting message as it is sent.
        /// </summary>
        /// <param name="kind">An invite/update, a cancellation of the whole meeting, or one to removed attendees.</param>
        /// <param name="until">The delivery time stamped on the appointment, or null.</param>
        /// <param name="queued">An earlier invite for this meeting is already waiting in the Outbox.</param>
        /// <param name="command">What Prodigo's own Send later form asked for, if this send came from it.</param>
        /// <param name="latestWaiting">The latest delivery time of this meeting's messages still in the Outbox, if any.</param>
        public static MeetingDecision Decide(MeetingMessageKind kind, DateTime? until, bool queued, PendingCommand command,
            DateTime now, DateTime? latestWaiting = null)
        {
            if (command == PendingCommand.SendNow)
                return new MeetingDecision { CleanUp = queued, ClearStamp = true, Reason = "send now requested" };

            if (!until.HasValue || until.Value <= now)
            {
                // The invite has gone, but an update for it may still be waiting for its
                // extra minute. A change sent now must still arrive after that one.
                if (until.HasValue && queued && latestWaiting.HasValue && latestWaiting.Value > now)
                    return new MeetingDecision
                    {
                        Defer = true,
                        DeliverAt = latestWaiting.Value + UpdateOffset,
                        CleanUp = true,
                        Reason = "change while an earlier update still waits",
                    };

                return new MeetingDecision { Reason = until.HasValue ? "delivery time has passed; normal send" : "not scheduled" };
            }

            if (!queued)
            {
                if (kind == MeetingMessageKind.Request)
                    return new MeetingDecision { Defer = true, DeliverAt = until.Value, MarkQueued = true, Reason = "initial invite" };

                // A cancellation for a meeting whose invite never went anywhere.
                return new MeetingDecision { Reason = "cancellation, nothing queued; normal send" };
            }

            if (kind == MeetingMessageKind.Request && command == PendingCommand.Exact)
                return new MeetingDecision { Defer = true, DeliverAt = until.Value, CleanUp = true, Reason = "rescheduled from Send later; replaces the waiting invite" };

            // An update or cancellation while the invite still waits. Delivered just
            // after it, so that if replacing the waiting invite fails, attendees still
            // get the invite first and this second - never the other way round.
            return new MeetingDecision
            {
                Defer = true,
                DeliverAt = until.Value + UpdateOffset,
                CleanUp = true,
                Reason = kind == MeetingMessageKind.Request ? "update while invite waits" : "cancellation while invite waits",
            };
        }

        /// <summary>
        /// Whether a cancellation is for the whole meeting or only for people removed
        /// from it. Outlook's cancellation to removed people goes only to people no
        /// longer on the meeting; a whole-meeting cancellation goes to the attendees
        /// themselves. So if anyone it goes to is still an attendee, the meeting itself
        /// is being cancelled. When the recipients are unknown, it counts as one to
        /// removed people, which deletes less.
        /// </summary>
        public static MeetingMessageKind CancellationKind(bool statusSaysCanceled,
            IEnumerable<string> cancelRecipients, IEnumerable<string> attendees)
        {
            if (statusSaysCanceled) return MeetingMessageKind.CancelWholeMeeting;
            if (cancelRecipients == null || attendees == null) return MeetingMessageKind.CancelRemovedAttendees;

            HashSet<string> still = WaitingMessage.RecipientSet(attendees);
            return WaitingMessage.RecipientSet(cancelRecipients).Overlaps(still)
                ? MeetingMessageKind.CancelWholeMeeting
                : MeetingMessageKind.CancelRemovedAttendees;
        }

        /// <summary>
        /// Decides which waiting messages to delete once a new message for the same
        /// meeting is safely in the Outbox. Returns null when there is no proof yet that
        /// the new message got there: deleting the old invite before then could leave
        /// attendees with nothing at all.
        /// </summary>
        /// <param name="waiting">Every Prodigo-tagged message for this meeting still in the Outbox.</param>
        /// <param name="batch">What was sent for this meeting in the send that triggered the cleanup.</param>
        /// <param name="proofFound">The newest message of the batch has been seen in the Outbox or Sent Items.</param>
        public static CleanupPlan PlanCleanup(IList<WaitingMessage> waiting, CleanupBatch batch, bool proofFound)
        {
            if (!proofFound) return null;

            var plan = new CleanupPlan();

            // A whole-meeting cancellation never comes in the same press of Send as an
            // update. If both are here, the "whole" one was misjudged and was really for
            // removed people: deleting everything would leave the rest with nothing.
            if (batch.WholeCancelGen != null && batch.RequestGen == null)
            {
                // The meeting was cancelled before anyone heard of it: nobody needs to hear
                // anything, the cancellation included.
                plan.First.AddRange(waiting.Select(w => w.Gen));
                return plan;
            }

            // An old invite is replaced only when this send accounts for everyone it went
            // to: each of them either gets the new invite, or was removed in this same
            // send. "Send updates only to added or deleted attendees", or forwarding the
            // meeting, sends a new invite to a few people - and deleting the original
            // then would leave everyone else with nothing. Unknown recipients (an
            // untagged message) count as "not covered".
            // With no new invite in the batch, only the removed people count: removing
            // everyone an old invite went to means nobody needs it any more.
            if (batch.RequestGen == null ? batch.RemovedCancelGens.Count > 0 : batch.RequestRecipients != null)
            {
                var covered = new HashSet<string>(
                    batch.RequestRecipients ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
                foreach (WaitingMessage w in waiting)
                    if (w.Recipients != null && (batch.RemovedCancelGens.Contains(w.Gen) || w.Gen == batch.WholeCancelGen))
                        covered.UnionWith(w.Recipients);

                plan.First.AddRange(waiting
                    .Where(w => w.Kind == MeetingMessageKind.Request && w.Gen != batch.RequestGen)
                    .Where(w => w.Recipients != null && w.Recipients.Count > 0 && w.Recipients.IsSubsetOf(covered))
                    .Select(w => w.Gen));
            }

            // A cancellation to removed attendees is dropped only if none of them is on
            // an invite that is still going out - otherwise they will get that invite
            // and need the cancellation that goes with it.
            List<WaitingMessage> remaining = waiting
                .Where(w => w.Kind == MeetingMessageKind.Request && !plan.First.Contains(w.Gen))
                .ToList();
            plan.Then.AddRange(waiting
                .Where(w => w.Kind == MeetingMessageKind.CancelRemovedAttendees && w.Recipients != null)
                .Where(w => !remaining.Any(r => r.Recipients == null || r.Recipients.Overlaps(w.Recipients)))
                .Select(w => w.Gen));
            return plan;
        }

        /// <summary>
        /// True when someone was removed from a meeting whose invite still waits, and
        /// Outlook sent them a cancellation but no update to the others ("Send updates
        /// only to added or deleted attendees", Outlook's default). The waiting invite
        /// still lists the removed people, so it cannot simply go. Prodigo then sends the
        /// update to the remaining attendees itself; with that in the batch, the old
        /// invite and the cancellation are both replaced, and the removed people get
        /// nothing at all.
        /// </summary>
        public static bool NeedsUpdateForRemoved(IList<WaitingMessage> waiting, CleanupBatch batch, CleanupPlan plan)
        {
            if (plan == null || batch.RequestGen != null || batch.WholeCancelGen != null || batch.RemovedCancelGens.Count == 0)
                return false;

            List<WaitingMessage> cancels = waiting
                .Where(w => batch.RemovedCancelGens.Contains(w.Gen) && w.Recipients != null && !plan.Then.Contains(w.Gen))
                .ToList();
            return waiting
                .Where(w => w.Kind == MeetingMessageKind.Request && w.Recipients != null && !plan.First.Contains(w.Gen))
                .Any(w => cancels.Any(c => c.Recipients.Overlaps(w.Recipients)));
        }
    }

    internal sealed class SendLaterPreset
    {
        private readonly Func<DateTime, DateTime> _resolve;

        public SendLaterPreset(string spec, string label, Func<DateTime, DateTime> resolve)
        {
            Spec = spec;
            Label = label;
            _resolve = resolve;
        }

        public string Spec { get; }

        /// <summary>"Tomorrow 08:00". The button adds the actual date next to it.</summary>
        public string Label { get; }

        /// <summary>The time this preset means right now, or null when that has already passed ("today 13:00" at 14:00).</summary>
        public DateTime? Resolve(DateTime now)
        {
            DateTime t = _resolve(now);
            return SendLaterTime.Validate(t, now) == null ? t : (DateTime?)null;
        }
    }

    internal enum MeetingMessageKind
    {
        Request,
        CancelWholeMeeting,
        CancelRemovedAttendees,
    }

    internal enum PendingCommand
    {
        None,
        /// <summary>Deliver at exactly the stamped time, replacing whatever waits.</summary>
        Exact,
        /// <summary>Deliver now, replacing whatever waits.</summary>
        SendNow,
    }

    internal sealed class MeetingDecision
    {
        public bool Defer { get; set; }
        public DateTime DeliverAt { get; set; }
        /// <summary>Remember on the appointment that its invite is now waiting in the Outbox.</summary>
        public bool MarkQueued { get; set; }
        /// <summary>Once this message is in the Outbox, remove the older ones it replaces.</summary>
        public bool CleanUp { get; set; }
        /// <summary>Remove the delivery-time stamp from the appointment afterwards.</summary>
        public bool ClearStamp { get; set; }
        public string Reason { get; set; }
    }

    internal sealed class WaitingMessage
    {
        public WaitingMessage(string gen, MeetingMessageKind kind, IEnumerable<string> recipients = null)
        {
            Gen = gen;
            Kind = kind;
            Recipients = recipients == null ? null : RecipientSet(recipients);
        }

        public string Gen { get; }
        public MeetingMessageKind Kind { get; }
        /// <summary>Lower-cased addresses, or null when not known.</summary>
        public HashSet<string> Recipients { get; }

        public static HashSet<string> RecipientSet(IEnumerable<string> addresses) =>
            new HashSet<string>(addresses.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => a.Trim()),
                StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>What one press of Send / Send update produced for one meeting.</summary>
    internal sealed class CleanupBatch
    {
        public string RequestGen { get; set; }
        /// <summary>Who the new invite or update in this batch goes to.</summary>
        public HashSet<string> RequestRecipients { get; set; }
        public string WholeCancelGen { get; set; }
        public List<string> RemovedCancelGens { get; } = new List<string>();
        /// <summary>Sent without delay (Send now), so it may already have left the Outbox.</summary>
        public bool Immediate { get; set; }
        /// <summary>The organizer's appointment, for sending the update to the remaining attendees.</summary>
        public string AppointmentEntryId { get; set; }

        /// <summary>The message whose arrival in the Outbox proves the send went through.</summary>
        public string ProofGen => WholeCancelGen ?? RequestGen ?? RemovedCancelGens.LastOrDefault();
    }

    internal sealed class CleanupPlan
    {
        public List<string> First { get; } = new List<string>();
        public List<string> Then { get; } = new List<string>();
    }
}
