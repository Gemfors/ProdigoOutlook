using System;
using System.Collections.Generic;
using ProdigoOutlook.Core;
using ProdigoOutlook.Features.SendLater;
using ProdigoOutlook.Features.Availability;
using System.Linq;

namespace ProdigoOutlook.LogicTests
{
    /// <summary>
    /// Plain checks, no test framework: each Check prints a line, and the exit code is
    /// the number of failures. Dates are fixed so the results never depend on today.
    /// Wednesday 4 November 2026 is the reference day.
    /// </summary>
    internal static class Program
    {
        private static int _failures;

        private static readonly DateTime Wed1000 = new DateTime(2026, 11, 4, 10, 0, 0);

        private static int Main()
        {
            Presets();
            Validation();
            MeetingDecisions();
            Cleanup();
            Stored();
            Availability();

            Console.WriteLine();
            Console.WriteLine(_failures == 0 ? "All checks passed." : _failures + " check(s) FAILED.");
            return _failures;
        }

        private static void Presets()
        {
            Section("Presets");

            SendLaterPreset hour = SendLaterTime.ParsePreset("+1h");
            Check("+1h label", hour?.Label, "In 1 hour");
            Check("+1h resolves", hour?.Resolve(Wed1000.AddSeconds(30)), new DateTime(2026, 11, 4, 11, 0, 0));
            Check("+30m label", SendLaterTime.ParsePreset("+30m")?.Label, "In 30 minutes");
            Check("+1h30m label", SendLaterTime.ParsePreset("+1h30m")?.Label, "In 1 h 30 min");

            Check("tomorrow 08:00", SendLaterTime.ParsePreset("tomorrow 08:00")?.Resolve(Wed1000), new DateTime(2026, 11, 5, 8, 0, 0));
            Check("imorgon 8.00 (Swedish, dot)", SendLaterTime.ParsePreset("imorgon 8.00")?.Resolve(Wed1000), new DateTime(2026, 11, 5, 8, 0, 0));

            DateTime fri = new DateTime(2026, 11, 6, 15, 0, 0);
            Check("next workday from Friday is Monday", SendLaterTime.ParsePreset("nextworkday 08:00")?.Resolve(fri), new DateTime(2026, 11, 9, 8, 0, 0));
            Check("next workday from Wednesday is Thursday", SendLaterTime.ParsePreset("nextworkday 08:00")?.Resolve(Wed1000), new DateTime(2026, 11, 5, 8, 0, 0));

            DateTime mon0700 = new DateTime(2026, 11, 9, 7, 0, 0);
            DateTime mon0900 = new DateTime(2026, 11, 9, 9, 0, 0);
            Check("monday 08:00 on Monday 07:00 is today", SendLaterTime.ParsePreset("monday 08:00")?.Resolve(mon0700), new DateTime(2026, 11, 9, 8, 0, 0));
            Check("monday 08:00 on Monday 09:00 is next week", SendLaterTime.ParsePreset("monday 08:00")?.Resolve(mon0900), new DateTime(2026, 11, 16, 8, 0, 0));
            Check("monday 08:00 from Wednesday", SendLaterTime.ParsePreset("monday 08:00")?.Resolve(Wed1000), new DateTime(2026, 11, 9, 8, 0, 0));
            Check("måndag works", SendLaterTime.ParsePreset("måndag 08:00")?.Resolve(Wed1000), new DateTime(2026, 11, 9, 8, 0, 0));
            Check("monday label", SendLaterTime.ParsePreset("monday 08:00")?.Label, "Monday 08:00");

            Check("today 13:00 before 13", SendLaterTime.ParsePreset("today 13:00")?.Resolve(Wed1000), new DateTime(2026, 11, 4, 13, 0, 0));
            Check("today 09:00 after 9 has passed", SendLaterTime.ParsePreset("today 09:00")?.Resolve(Wed1000), (DateTime?)null);

            Check("nonsense is rejected", SendLaterTime.ParsePreset("soon") == null, true);
            Check("25:00 is rejected", SendLaterTime.ParsePreset("tomorrow 25:00") == null, true);
            Check("+0h is rejected", SendLaterTime.ParsePreset("+0h") == null, true);
            Check("empty is rejected", SendLaterTime.ParsePreset("  ") == null, true);
        }

        private static void Validation()
        {
            Section("Validation");
            Check("past rejected", SendLaterTime.Validate(Wed1000.AddMinutes(-5), Wed1000) != null, true);
            Check("now rejected", SendLaterTime.Validate(Wed1000, Wed1000) != null, true);
            Check("30 s ahead rejected", SendLaterTime.Validate(Wed1000.AddSeconds(30), Wed1000) != null, true);
            Check("2 min ahead accepted", SendLaterTime.Validate(Wed1000.AddMinutes(2), Wed1000), (string)null);
        }

        private static void MeetingDecisions()
        {
            Section("Meeting decisions");
            DateTime until = new DateTime(2026, 11, 5, 8, 0, 0);

            MeetingDecision d = SendLaterTime.Decide(MeetingMessageKind.Request, until, false, PendingCommand.Exact, Wed1000);
            Check("initial invite deferred to stamp", d.Defer && d.DeliverAt == until, true);
            Check("initial invite marks queued", d.MarkQueued, true);
            Check("initial invite needs no cleanup", d.CleanUp, false);

            d = SendLaterTime.Decide(MeetingMessageKind.Request, until, false, PendingCommand.None, Wed1000);
            Check("initial via Outlook's own Send button", d.Defer && d.DeliverAt == until, true);

            d = SendLaterTime.Decide(MeetingMessageKind.Request, until, true, PendingCommand.None, Wed1000);
            Check("update while waiting: stamp + 1 min", d.Defer && d.DeliverAt == until.AddMinutes(1), true);
            Check("update while waiting: cleanup", d.CleanUp, true);

            d = SendLaterTime.Decide(MeetingMessageKind.Request, until, true, PendingCommand.Exact, Wed1000);
            Check("reschedule from form: exact stamp", d.Defer && d.DeliverAt == until, true);
            Check("reschedule from form: cleanup", d.CleanUp, true);

            d = SendLaterTime.Decide(MeetingMessageKind.CancelWholeMeeting, until, true, PendingCommand.None, Wed1000);
            Check("cancellation while waiting: stamp + 1 min", d.Defer && d.DeliverAt == until.AddMinutes(1), true);
            Check("cancellation while waiting: cleanup", d.CleanUp, true);

            d = SendLaterTime.Decide(MeetingMessageKind.CancelRemovedAttendees, until, true, PendingCommand.None, Wed1000);
            Check("removed attendees while waiting: stamp + 1 min", d.Defer && d.DeliverAt == until.AddMinutes(1), true);

            d = SendLaterTime.Decide(MeetingMessageKind.CancelWholeMeeting, until, false, PendingCommand.None, Wed1000);
            Check("cancellation, nothing queued: normal send", d.Defer || d.CleanUp, false);

            d = SendLaterTime.Decide(MeetingMessageKind.Request, until, true, PendingCommand.SendNow, Wed1000);
            Check("send now: not deferred", d.Defer, false);
            Check("send now: cleanup and clear", d.CleanUp && d.ClearStamp, true);

            d = SendLaterTime.Decide(MeetingMessageKind.Request, Wed1000.AddHours(-1), true, PendingCommand.None, Wed1000);
            Check("stamp passed: normal send", d.Defer || d.CleanUp, false);

            d = SendLaterTime.Decide(MeetingMessageKind.Request, null, false, PendingCommand.None, Wed1000);
            Check("no stamp: normal send", d.Defer || d.CleanUp, false);

            DateTime justPassed = Wed1000.AddSeconds(-20);
            DateTime updateWaits = justPassed.AddMinutes(1);
            d = SendLaterTime.Decide(MeetingMessageKind.Request, justPassed, true, PendingCommand.None, Wed1000, updateWaits);
            Check("change in the minute after delivery: after the waiting update", d.Defer && d.DeliverAt == updateWaits.AddMinutes(1), true);
            Check("change in the minute after delivery: cleanup", d.CleanUp, true);
        }

        private static void Cleanup()
        {
            Section("Replacing waiting invites");

            // The original invite went to a, b and c. The update removes c.
            var waiting = new List<WaitingMessage>
            {
                new WaitingMessage("old", MeetingMessageKind.Request, new[] { "a@x", "b@x", "c@x" }),
                new WaitingMessage("new", MeetingMessageKind.Request, new[] { "a@x", "b@x" }),
                new WaitingMessage("rm", MeetingMessageKind.CancelRemovedAttendees, new[] { "c@x" }),
            };

            var update = new CleanupBatch { RequestGen = "new", RequestRecipients = WaitingMessage.RecipientSet(new[] { "A@x", "b@x" }) };
            update.RemovedCancelGens.Add("rm");

            Check("no proof, no deletion", SendLaterTime.PlanCleanup(waiting, update, false) == null, true);

            // a and b get the new invite and c was removed in the same send, so the old
            // invite is fully accounted for: it goes, and so does c's cancellation.
            CleanupPlan plan = SendLaterTime.PlanCleanup(waiting, update, true);
            Check("removing c: old invite replaced", string.Join(",", plan.First), "old");
            Check("removing c: c gets nothing at all", string.Join(",", plan.Then), "rm");

            // Same meeting, a full update to everyone still on it plus a new person d.
            var full = new List<WaitingMessage>
            {
                new WaitingMessage("old", MeetingMessageKind.Request, new[] { "a@x", "b@x" }),
                new WaitingMessage("new", MeetingMessageKind.Request, new[] { "a@x", "b@x", "d@x" }),
            };
            var fullBatch = new CleanupBatch { RequestGen = "new", RequestRecipients = WaitingMessage.RecipientSet(new[] { "a@x", "b@x", "d@x" }) };
            plan = SendLaterTime.PlanCleanup(full, fullBatch, true);
            Check("full update replaces the old invite", string.Join(",", plan.First), "old");

            // "Send updates only to added attendees": the new invite goes to d alone.
            var addedOnly = new CleanupBatch { RequestGen = "new", RequestRecipients = WaitingMessage.RecipientSet(new[] { "d@x" }) };
            var addedWaiting = new List<WaitingMessage>
            {
                new WaitingMessage("old", MeetingMessageKind.Request, new[] { "a@x", "b@x" }),
                new WaitingMessage("new", MeetingMessageKind.Request, new[] { "d@x" }),
            };
            plan = SendLaterTime.PlanCleanup(addedWaiting, addedOnly, true);
            Check("invite to added people only keeps the original", plan.First.Count, 0);

            // An old invite whose recipients are unknown is never deleted.
            var unknown = new List<WaitingMessage>
            {
                new WaitingMessage("old", MeetingMessageKind.Request),
                new WaitingMessage("new", MeetingMessageKind.Request, new[] { "a@x" }),
            };
            plan = SendLaterTime.PlanCleanup(unknown, new CleanupBatch { RequestGen = "new", RequestRecipients = WaitingMessage.RecipientSet(new[] { "a@x" }) }, true);
            Check("unknown recipients are not replaced", plan.First.Count, 0);

            // Removing c, where the old invite had been replaced by a full update earlier:
            // only the update to a and b is waiting, so c's cancellation can go.
            var afterReplace = new List<WaitingMessage>
            {
                new WaitingMessage("new", MeetingMessageKind.Request, new[] { "a@x", "b@x" }),
                new WaitingMessage("rm", MeetingMessageKind.CancelRemovedAttendees, new[] { "c@x" }),
            };
            plan = SendLaterTime.PlanCleanup(afterReplace, update, true);
            Check("cancellation dropped when no waiting invite includes c", string.Join(",", plan.Then), "rm");

            var cancel = new CleanupBatch { WholeCancelGen = "cx" };
            var waitingWithCancel = new List<WaitingMessage>(waiting) { new WaitingMessage("cx", MeetingMessageKind.CancelWholeMeeting) };
            plan = SendLaterTime.PlanCleanup(waitingWithCancel, cancel, true);
            Check("whole cancellation deletes everything", plan.First.Count == 4 && plan.Then.Count == 0, true);

            var removedOnly = new CleanupBatch();
            removedOnly.RemovedCancelGens.Add("rm");
            Check("proof is the cancellation when that is all there is", removedOnly.ProofGen, "rm");
            plan = SendLaterTime.PlanCleanup(waiting, removedOnly, true);
            Check("removal with no new invite deletes no invite", plan.First.Count, 0);
            Check("...and keeps the cancellation", plan.Then.Count, 0);

            // A "whole" cancellation sent in the same press as an update was misjudged:
            // it must not wipe out the update everyone else is waiting for.
            var mixed = new CleanupBatch { RequestGen = "new", RequestRecipients = WaitingMessage.RecipientSet(new[] { "a@x", "b@x" }), WholeCancelGen = "cx" };
            var mixedWaiting = new List<WaitingMessage>
            {
                new WaitingMessage("old", MeetingMessageKind.Request, new[] { "a@x", "b@x", "c@x" }),
                new WaitingMessage("new", MeetingMessageKind.Request, new[] { "a@x", "b@x" }),
                new WaitingMessage("cx", MeetingMessageKind.CancelWholeMeeting, new[] { "c@x" }),
            };
            plan = SendLaterTime.PlanCleanup(mixedWaiting, mixed, true);
            Check("misjudged whole cancel with an update keeps the update", plan.First.Contains("new"), false);
            Check("...and still replaces the old invite", string.Join(",", plan.First), "old");

            // Test C, 5 Nov 2026: invite waits for a and t; t is removed with "send only to
            // added or deleted attendees", so only a cancellation to t is sent.
            var testC = new List<WaitingMessage>
            {
                new WaitingMessage("orig", MeetingMessageKind.Request, new[] { "a@x", "t@x" }),
                new WaitingMessage("rmT", MeetingMessageKind.CancelRemovedAttendees, new[] { "t@x" }),
            };
            var cancelOnly = new CleanupBatch();
            cancelOnly.RemovedCancelGens.Add("rmT");
            plan = SendLaterTime.PlanCleanup(testC, cancelOnly, true);
            Check("removal only: original invite kept for now", plan.First.Count + plan.Then.Count, 0);

            // A full update to a in the same batch as the cancellation (choosing "all attendees").
            var afterAuto = new List<WaitingMessage>(testC) { new WaitingMessage("auto", MeetingMessageKind.Request, new[] { "a@x" }) };
            var carried = new CleanupBatch { RequestGen = "auto", RequestRecipients = WaitingMessage.RecipientSet(new[] { "a@x" }) };
            carried.RemovedCancelGens.Add("rmT");
            plan = SendLaterTime.PlanCleanup(afterAuto, carried, true);
            Check("after the update: original invite replaced", string.Join(",", plan.First), "orig");
            Check("after the update: t gets nothing", string.Join(",", plan.Then), "rmT");

            // Removing the only attendee: nobody is left to need the invite.
            var onlyOne = new List<WaitingMessage>
            {
                new WaitingMessage("orig", MeetingMessageKind.Request, new[] { "t@x" }),
                new WaitingMessage("rmT", MeetingMessageKind.CancelRemovedAttendees, new[] { "t@x" }),
            };
            plan = SendLaterTime.PlanCleanup(onlyOne, cancelOnly, true);
            Check("removing the only attendee drops both", string.Join(",", plan.First) + "|" + string.Join(",", plan.Then), "orig|rmT");


            Section("Which cancellation is it");
            // 5 Nov 2026: Outlook's Cancel Meeting arrived with the status not yet cancelled.
            Check("cancellation to an attendee is the whole meeting",
                SendLaterTime.CancellationKind(false, new[] { "Gmail@x" }, new[] { "me@x", "gmail@x" }), MeetingMessageKind.CancelWholeMeeting);
            Check("cancellation to someone no longer on it is for removed people",
                SendLaterTime.CancellationKind(false, new[] { "c@x" }, new[] { "me@x", "a@x" }), MeetingMessageKind.CancelRemovedAttendees);
            Check("status says cancelled wins",
                SendLaterTime.CancellationKind(true, new[] { "c@x" }, new[] { "a@x" }), MeetingMessageKind.CancelWholeMeeting);
            Check("unknown recipients count as removed people",
                SendLaterTime.CancellationKind(false, null, new[] { "a@x" }), MeetingMessageKind.CancelRemovedAttendees);
        }

        private static void Stored()
        {
            Section("Stored times");
            DateTime local = new DateTime(2026, 11, 5, 8, 0, 0, DateTimeKind.Local);
            Check("round trip", TimeHelpers.FromStored(TimeHelpers.ToStored(local)), local);
            Check("stored is UTC text", TimeHelpers.ToStored(local).EndsWith("Z"), true);
            Check("garbage reads as null", TimeHelpers.FromStored("tomorrow"), (DateTime?)null);
            Check("4501 is no date", TimeHelpers.IsNoDate(TimeHelpers.OutlookNoDate), true);
        }

        private static void Availability()
        {
            Section("Availability: period");
            DateTime first, last;
            AvailabilityText.PeriodDays(AvailabilityPeriod.ThisWeek, Wed1000, out first, out last);
            Check("this week: from today to Sunday", first.ToString("dd") + "-" + last.ToString("dd"), "04-08");
            AvailabilityText.PeriodDays(AvailabilityPeriod.NextWeek, Wed1000, out first, out last);
            Check("next week: Monday to Sunday", first.ToString("dd") + "-" + last.ToString("dd"), "09-15");
            AvailabilityText.PeriodDays(AvailabilityPeriod.NextTwoWeeks, Wed1000, out first, out last);
            Check("next 2 weeks: today plus 13 days", first.ToString("dd") + "-" + last.ToString("dd"), "04-17");

            Section("Availability: free time");
            DateTime wed = new DateTime(2026, 11, 4);
            DateTime now = wed.AddHours(10).AddMinutes(10);
            var busy = new List<TimeRange>
            {
                new TimeRange(wed.AddHours(11), wed.AddHours(12)),
                new TimeRange(wed.AddHours(13).AddMinutes(15), wed.AddHours(14)),
            };
            var rules = new AvailabilityRules();
            Check("earliest is the next half hour; edges rounded inwards",
                Ranges(AvailabilityText.FreeRanges(busy, rules, wed, wed, now, TimeSpan.FromMinutes(30))),
                "10:30-11:00 12:00-13:00 14:00-17:00");
            Check("gaps too short for the meeting are left out",
                Ranges(AvailabilityText.FreeRanges(busy, rules, wed, wed, now, TimeSpan.FromMinutes(60))),
                "12:00-13:00 14:00-17:00");
            Check("buffer keeps time clear around meetings",
                Ranges(AvailabilityText.FreeRanges(busy, new AvailabilityRules { Buffer = TimeSpan.FromMinutes(15) }, wed, wed, now, TimeSpan.FromMinutes(30))),
                "12:30-13:00 14:30-17:00");

            var overlapping = new List<TimeRange>
            {
                new TimeRange(wed.AddHours(9), wed.AddHours(12)),
                new TimeRange(wed.AddHours(10), wed.AddHours(11)),
                new TimeRange(wed.AddHours(16).AddMinutes(45), wed.AddHours(18)),
            };
            Check("overlapping meetings, and one running past the end of the day",
                Ranges(AvailabilityText.FreeRanges(overlapping, rules, wed, wed, wed.AddHours(7), TimeSpan.FromMinutes(30))),
                "08:00-09:00 12:00-16:30");

            Check("after working hours today, nothing today",
                AvailabilityText.FreeRanges(busy, rules, wed, wed, wed.AddHours(17).AddMinutes(5), TimeSpan.FromMinutes(30)).Count, 0);

            DateTime sat = new DateTime(2026, 11, 7);
            Check("weekend skipped",
                AvailabilityText.FreeRanges(new List<TimeRange>(), rules, sat, sat.AddDays(1), wed, TimeSpan.FromMinutes(30)).Count, 0);

            var allDay = new List<TimeRange> { new TimeRange(wed.AddDays(1), wed.AddDays(2)) };
            Check("an all-day busy event blocks the day",
                AvailabilityText.FreeRanges(allDay, rules, wed.AddDays(1), wed.AddDays(1), wed, TimeSpan.FromMinutes(30)).Count, 0);

            Section("Availability: picking and text");
            var many = new List<TimeRange>
            {
                new TimeRange(wed.AddHours(8), wed.AddHours(9)),
                new TimeRange(wed.AddHours(10), wed.AddHours(11)),
                new TimeRange(wed.AddHours(13), wed.AddHours(14)),
                new TimeRange(wed.AddHours(15), wed.AddHours(16)),
                new TimeRange(wed.AddDays(1).AddHours(9), wed.AddDays(1).AddHours(10)),
                new TimeRange(wed.AddDays(2).AddHours(9), wed.AddDays(2).AddHours(10)),
                new TimeRange(wed.AddDays(2).AddHours(14), wed.AddDays(2).AddHours(15)),
            };
            List<TimeRange> picked = AvailabilityText.Pick(many);
            Check("at most 5, spread over the days, in time order",
                string.Join(" ", picked.Select(r => r.Start.ToString("ddd HH", System.Globalization.CultureInfo.InvariantCulture))),
                "Wed 08 Wed 10 Thu 09 Fri 09 Fri 14");

            var two = new List<TimeRange>
            {
                new TimeRange(new DateTime(2026, 11, 10, 10, 0, 0), new DateTime(2026, 11, 10, 12, 0, 0)),
                new TimeRange(new DateTime(2026, 11, 11, 14, 0, 0), new DateTime(2026, 11, 11, 17, 0, 0)),
            };
            Check("English list",
                AvailabilityText.Format(two, AvailabilityFormat.List, AvailabilityLanguage.English, "I'm available at the following times:"),
                "I'm available at the following times:\n• Tue 10 Nov, 10:00–12:00\n• Wed 11 Nov, 14:00–17:00");
            Check("list for real Word bullets: no bullet characters",
                AvailabilityText.Format(two, AvailabilityFormat.List, AvailabilityLanguage.English, "", ""),
                "Tue 10 Nov, 10:00–12:00\nWed 11 Nov, 14:00–17:00");
            Check("Swedish sentence",
                AvailabilityText.Format(two, AvailabilityFormat.Sentence, AvailabilityLanguage.Swedish, "Jag kan"),
                "Jag kan tis 10 nov 10:00–12:00 eller ons 11 nov 14:00–17:00.");
            Check("nothing free gives no text",
                AvailabilityText.Format(new List<TimeRange>(), AvailabilityFormat.List, AvailabilityLanguage.English, "x"), null);

            Section("Availability: settings text");
            Check("time 8.30", AvailabilityText.ParseTime("8.30"), (TimeSpan?)new TimeSpan(8, 30, 0));
            Check("time nonsense", AvailabilityText.ParseTime("eight"), (TimeSpan?)null);
            Check("days in Swedish", AvailabilityText.DaysText(AvailabilityText.ParseDays("mån, tis, onsdag")), "mon,tue,wed");
            Check("bad day refused", AvailabilityText.ParseDays("mon,funday"), (HashSet<DayOfWeek>)null);
            AvailabilityRules fallback = AvailabilityText.Rules("17:00", "08:00", "nonsense", 15, "13:00", "12:00");
            Check("lunch the wrong way round means no lunch break", fallback.LunchStart.HasValue, false);
            Check("empty lunch means no lunch break", AvailabilityText.Rules("08:30", "17:00", "mon", 0, "", "").LunchStart.HasValue, false);

            DateTime thu = new DateTime(2026, 11, 5);
            AvailabilityRules olle = AvailabilityText.Rules("08:30", "17:00", "mon,tue,wed,thu,fri", 0, "12:00", "13:00");
            Check("Olle's day: from 08:30, never over lunch",
                Ranges(AvailabilityText.FreeRanges(new List<TimeRange>(), olle, thu, thu, thu, TimeSpan.FromMinutes(30))),
                "08:30-12:00 13:00-17:00");
            AvailabilityRules olleBuffer = AvailabilityText.Rules("08:30", "17:00", "mon,tue,wed,thu,fri", 15, "12:00", "13:00");
            Check("the buffer is not added around lunch",
                Ranges(AvailabilityText.FreeRanges(new List<TimeRange>(), olleBuffer, thu, thu, thu, TimeSpan.FromMinutes(30))),
                "08:30-12:00 13:00-17:00");
            Check("unreadable settings fall back to 08-17 Mon-Fri",
                fallback.WorkStart + " " + fallback.WorkEnd + " " + AvailabilityText.DaysText(fallback.WorkDays) + " " + fallback.Buffer.TotalMinutes,
                "08:00:00 17:00:00 mon,tue,wed,thu,fri 15");

            Section("Availability: language");
            Check("Swedish email", AvailabilityText.DetectLanguage("Hej Olle! Tack för senast, kan vi ses och prata om det?"), AvailabilityLanguage.Swedish);
            Check("English email", AvailabilityText.DetectLanguage("Hi Olle, thanks for the call. Would you have time to meet and go through the plan?"), AvailabilityLanguage.English);
            Check("too little text: undecided", AvailabilityText.DetectLanguage("OK"), (AvailabilityLanguage?)null);

            // A reply as it really looks in Olle's English Outlook: English header lines
            // above the quoted Swedish email, and an English disclaimer below it.
            const string swedishReply =
                "\r\n\r\nMvh\r\nOlle Gemfors | Group Strategy Manager | HiQ\r\n\r\n" +
                "From: Anna Svensson <anna.svensson@example.se>\r\n" +
                "Sent: Monday, November 9, 2026 4:12 PM\r\n" +
                "To: Olle Gemfors <olle.gemfors@hiq.se>\r\n" +
                "Cc: Per Nilsson <per.nilsson@example.se>\r\n" +
                "Subject: RE: Möte om strategin\r\n\r\n" +
                "Hej Olle!\r\n\r\nVi skulle vilja boka in ett möte för att gå igenom planen inför nästa år. " +
                "När passar det för dig?\r\n\r\nVänliga hälsningar\r\nAnna\r\n\r\n" +
                "This e-mail and any attachments are confidential and may be legally privileged. If you are not " +
                "the intended recipient, please notify the sender immediately and delete it from your system. " +
                "Any use of the information in this e-mail by anyone other than the intended recipient is " +
                "prohibited, and we do not accept any liability for the contents of this message or for any " +
                "changes made to it after it was sent. Please consider the environment before printing.";
            int svCount, enCount;
            AvailabilityLanguage? detected = AvailabilityText.DetectLanguage(swedishReply, out svCount, out enCount);
            Check("Swedish reply with English headers and disclaimer", detected, AvailabilityLanguage.Swedish);
            if (detected != AvailabilityLanguage.Swedish)
                Console.WriteLine("        sv=" + svCount + " en=" + enCount + " sample: " + AvailabilityText.LanguageSample(swedishReply));

            const string englishReply =
                "\r\n\r\nMvh\r\nOlle\r\n\r\nFrom: John Smith <john@example.com>\r\nSent: Monday\r\nTo: Olle\r\n" +
                "Subject: Strategy meeting\r\n\r\nHi Olle, would you have time next week to go through the plan " +
                "with me and the team? Thanks, John";
            Check("English reply with Olle's Swedish sign-off", AvailabilityText.DetectLanguage(englishReply), AvailabilityLanguage.English);
        }

        private static string Ranges(IEnumerable<TimeRange> ranges) =>
            string.Join(" ", ranges.Select(r => r.Start.ToString("HH:mm") + "-" + r.End.ToString("HH:mm")));

        private static CleanupBatch removedOnlyFor(string gen)
        {
            var b = new CleanupBatch();
            b.RemovedCancelGens.Add(gen);
            return b;
        }

        private static void Section(string name)
        {
            Console.WriteLine();
            Console.WriteLine("== " + name);
        }

        private static void Check<T>(string name, T actual, T expected)
        {
            bool ok = EqualityComparer<T>.Default.Equals(actual, expected);
            if (!ok) _failures++;
            Console.WriteLine((ok ? "  ok    " : "  FAIL  ") + name + (ok ? "" : "  (got " + Show(actual) + ", expected " + Show(expected) + ")"));
        }

        private static string Show(object o) => o == null ? "null" : o.ToString();
    }
}
