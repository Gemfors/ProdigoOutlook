using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProdigoOutlook.Features.Availability
{
    /// <summary>
    /// Insert availability without Outlook: busy times in, free time ranges out, and
    /// ranges turned into the text that goes into the email. No Outlook types here, so
    /// all of it is tested in test\LogicTests. AvailabilityCalendar reads the busy
    /// times; AvailabilityForm shows the result.
    ///
    /// Agreed with Olle on 10 November 2026:
    /// - Free = within working hours on working days, outside anything not marked Free,
    ///   in a gap the meeting fits in. Only his own calendar.
    /// - Earliest = the next whole or half hour from now. Everything is rounded to half
    ///   hours, inwards, so a range never reaches into a meeting.
    /// - Buffer around meetings: none by default, a setting.
    /// - Changed the same day at Olle's request: working day from 08:30, and lunch
    ///   12:00-13:00 never offered (a setting; empty means no lunch break).
    /// - At most 5 ranges; a list (default) or a sentence; Swedish or English.
    /// </summary>
    internal static class AvailabilityText
    {
        public const int MaxRanges = 5;
        public static readonly TimeSpan Step = TimeSpan.FromMinutes(30);

        public static readonly int[] Lengths = { 30, 45, 60, 90, 120 };

        // ---------------------------------------------------------------
        // Settings text
        // ---------------------------------------------------------------

        /// <summary>"08:00" or "8" or "8.30"; null when not a time of day.</summary>
        public static TimeSpan? ParseTime(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            string[] parts = text.Trim().Replace('.', ':').Split(':');
            int h, m = 0;
            if (parts.Length > 2 || !int.TryParse(parts[0], out h)) return null;
            if (parts.Length == 2 && !int.TryParse(parts[1], out m)) return null;
            if (h < 0 || h > 24 || m < 0 || m > 59 || (h == 24 && m > 0)) return null;
            return new TimeSpan(h, m, 0);
        }

        private static readonly Dictionary<string, DayOfWeek> DayNames = new Dictionary<string, DayOfWeek>(StringComparer.OrdinalIgnoreCase)
        {
            { "mon", DayOfWeek.Monday }, { "tue", DayOfWeek.Tuesday }, { "wed", DayOfWeek.Wednesday },
            { "thu", DayOfWeek.Thursday }, { "fri", DayOfWeek.Friday }, { "sat", DayOfWeek.Saturday }, { "sun", DayOfWeek.Sunday },
            { "mån", DayOfWeek.Monday }, { "tis", DayOfWeek.Tuesday }, { "ons", DayOfWeek.Wednesday },
            { "tor", DayOfWeek.Thursday }, { "fre", DayOfWeek.Friday }, { "lör", DayOfWeek.Saturday }, { "sön", DayOfWeek.Sunday },
        };

        /// <summary>"mon,tue,wed" (or Swedish "mån,tis"); null when any part is not a day.</summary>
        public static HashSet<DayOfWeek> ParseDays(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var days = new HashSet<DayOfWeek>();
            foreach (string raw in text.Split(new[] { ',', ' ', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string key = raw.Trim();
                if (key.Length > 3) key = key.Substring(0, 3);
                DayOfWeek d;
                if (!DayNames.TryGetValue(key, out d)) return null;
                days.Add(d);
            }
            return days.Count > 0 ? days : null;
        }

        public static string DaysText(IEnumerable<DayOfWeek> days)
        {
            string[] names = { "sun", "mon", "tue", "wed", "thu", "fri", "sat" };
            return string.Join(",", days.OrderBy(d => ((int)d + 6) % 7).Select(d => names[(int)d]));
        }

        /// <summary>
        /// Rules from the settings text. Anything unreadable falls back to the agreed
        /// default (08:00-17:00, Monday to Friday), never to "no working hours".
        /// </summary>
        public static AvailabilityRules Rules(string workStart, string workEnd, string workDays, int bufferMinutes,
            string lunchStart, string lunchEnd)
        {
            var rules = new AvailabilityRules();
            TimeSpan? s = ParseTime(workStart), e = ParseTime(workEnd);
            if (s.HasValue && e.HasValue && e.Value > s.Value)
            {
                rules.WorkStart = s.Value;
                rules.WorkEnd = e.Value;
            }
            HashSet<DayOfWeek> days = ParseDays(workDays);
            if (days != null) rules.WorkDays = days;
            rules.Buffer = TimeSpan.FromMinutes(Math.Max(0, Math.Min(240, bufferMinutes)));
            TimeSpan? ls = ParseTime(lunchStart), le = ParseTime(lunchEnd);
            if (ls.HasValue && le.HasValue && le.Value > ls.Value)
            {
                rules.LunchStart = ls;
                rules.LunchEnd = le;
            }
            return rules;
        }

        // ---------------------------------------------------------------
        // Period
        // ---------------------------------------------------------------

        /// <summary>The days a period covers, from the first to the last, inclusive.</summary>
        public static void PeriodDays(AvailabilityPeriod period, DateTime now, out DateTime first, out DateTime last)
        {
            DateTime today = now.Date;
            int sinceMonday = ((int)today.DayOfWeek + 6) % 7;
            DateTime monday = today.AddDays(-sinceMonday);
            switch (period)
            {
                case AvailabilityPeriod.NextWeek:
                    first = monday.AddDays(7);
                    last = monday.AddDays(13);
                    break;
                case AvailabilityPeriod.NextTwoWeeks:
                    first = today;
                    last = today.AddDays(13);
                    break;
                default:
                    first = today;
                    last = monday.AddDays(6);
                    break;
            }
        }

        // ---------------------------------------------------------------
        // Free time
        // ---------------------------------------------------------------

        /// <summary>
        /// Every free range in the period that the meeting fits in, in time order.
        /// </summary>
        public static List<TimeRange> FreeRanges(IEnumerable<TimeRange> busy, AvailabilityRules rules,
            DateTime firstDay, DateTime lastDay, DateTime now, TimeSpan length)
        {
            List<TimeRange> blocked = busy
                .Select(b => new TimeRange(b.Start - rules.Buffer, b.End + rules.Buffer))
                .Where(b => b.End > b.Start)
                .ToList();

            // Lunch is blocked like a meeting, but without the buffer: it is not a meeting.
            if (rules.LunchStart.HasValue && rules.LunchEnd.HasValue && rules.LunchEnd.Value > rules.LunchStart.Value)
                for (DateTime day = firstDay.Date; day <= lastDay.Date; day = day.AddDays(1))
                    blocked.Add(new TimeRange(day + rules.LunchStart.Value, day + rules.LunchEnd.Value));

            blocked = blocked.OrderBy(b => b.Start).ToList();

            DateTime earliest = RoundUp(now);
            var result = new List<TimeRange>();

            for (DateTime day = firstDay.Date; day <= lastDay.Date; day = day.AddDays(1))
            {
                if (!rules.WorkDays.Contains(day.DayOfWeek)) continue;

                DateTime start = day + rules.WorkStart;
                DateTime end = day + rules.WorkEnd;
                if (start < earliest) start = earliest;
                if (end <= start) continue;

                DateTime cursor = start;
                foreach (TimeRange b in blocked)
                {
                    if (b.End <= cursor) continue;
                    if (b.Start >= end) break;
                    Add(result, cursor, b.Start, length);
                    if (b.End > cursor) cursor = b.End;
                    if (cursor >= end) break;
                }
                Add(result, cursor, end, length);
            }
            return result;
        }

        private static void Add(List<TimeRange> result, DateTime start, DateTime end, TimeSpan length)
        {
            DateTime s = RoundUp(start), e = RoundDown(end);
            if (e - s >= length) result.Add(new TimeRange(s, e));
        }

        public static DateTime RoundUp(DateTime t)
        {
            long ticks = Step.Ticks;
            long rem = t.Ticks % ticks;
            return rem == 0 ? t : new DateTime(t.Ticks - rem + ticks, t.Kind);
        }

        public static DateTime RoundDown(DateTime t) => new DateTime(t.Ticks - t.Ticks % Step.Ticks, t.Kind);

        /// <summary>
        /// At most <paramref name="max"/> ranges, spread over the days rather than all
        /// from the first one: the first free range of each day, then each day's second,
        /// and so on. Returned in time order.
        /// </summary>
        public static List<TimeRange> Pick(IList<TimeRange> free, int max = MaxRanges)
        {
            List<List<TimeRange>> byDay = free.GroupBy(r => r.Start.Date).OrderBy(g => g.Key)
                .Select(g => g.OrderBy(r => r.Start).ToList()).ToList();
            var picked = new List<TimeRange>();
            for (int round = 0; picked.Count < max; round++)
            {
                bool any = false;
                foreach (List<TimeRange> day in byDay)
                {
                    if (round >= day.Count) continue;
                    any = true;
                    picked.Add(day[round]);
                    if (picked.Count == max) break;
                }
                if (!any) break;
            }
            return picked.OrderBy(r => r.Start).ToList();
        }

        // ---------------------------------------------------------------
        // Text
        // ---------------------------------------------------------------

        private static readonly string[] SvDays = { "sön", "mån", "tis", "ons", "tor", "fre", "lör" };
        private static readonly string[] EnDays = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        private static readonly string[] SvMonths = { "jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec" };
        private static readonly string[] EnMonths = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

        /// <summary>"tis 10 nov" / "Tue 10 Nov".</summary>
        public static string Day(DateTime d, AvailabilityLanguage lang) => lang == AvailabilityLanguage.Swedish
            ? SvDays[(int)d.DayOfWeek] + " " + d.Day + " " + SvMonths[d.Month - 1]
            : EnDays[(int)d.DayOfWeek] + " " + d.Day + " " + EnMonths[d.Month - 1];

        public static string Range(TimeRange r) => r.Start.ToString("HH:mm") + "–" + r.End.ToString("HH:mm");

        /// <summary>
        /// The text to insert. Lines are separated by "\n"; the caller turns them into
        /// paragraphs. Null when there is nothing to offer.
        /// </summary>
        /// <param name="bullet">
        /// Put before each list line: "• " for a plain-text email and the preview, ""
        /// when the lines get real Word bullets after inserting.
        /// </param>
        public static string Format(IList<TimeRange> ranges, AvailabilityFormat format, AvailabilityLanguage lang, string opening,
            string bullet = "• ")
        {
            if (ranges == null || ranges.Count == 0) return null;
            opening = (opening ?? string.Empty).Trim();

            if (format == AvailabilityFormat.List)
            {
                var sb = new StringBuilder();
                if (opening.Length > 0) sb.Append(opening).Append('\n');
                for (int i = 0; i < ranges.Count; i++)
                {
                    if (i > 0) sb.Append('\n');
                    sb.Append(bullet ?? string.Empty).Append(Day(ranges[i].Start, lang)).Append(", ").Append(Range(ranges[i]));
                }
                return sb.ToString();
            }

            List<string> parts = ranges.Select(r => Day(r.Start, lang) + " " + Range(r)).ToList();
            string or = lang == AvailabilityLanguage.Swedish ? " eller " : " or ";
            string joined = parts.Count == 1
                ? parts[0]
                : string.Join(", ", parts.Take(parts.Count - 1)) + or + parts[parts.Count - 1];
            return (opening.Length > 0 ? opening + " " : string.Empty) + joined + ".";
        }

        // ---------------------------------------------------------------
        // Language
        // ---------------------------------------------------------------

        private static readonly HashSet<string> SvWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "och", "att", "jag", "det", "är", "som", "för", "med", "inte", "på", "har", "vi", "kan", "du",
            "hej", "tack", "mvh", "hälsningar", "vänliga", "skulle", "också", "eller", "till", "från", "om",
        };

        private static readonly HashSet<string> EnWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "and", "the", "to", "is", "that", "for", "with", "not", "on", "have", "we", "can", "you",
            "hi", "hello", "thanks", "regards", "best", "would", "also", "or", "from", "of", "dear",
        };

        private const int DecisiveLead = 5;

        private static readonly string[] HeaderLabels =
        {
            "from:", "sent:", "to:", "cc:", "bcc:", "subject:", "date:",
            "från:", "skickat:", "till:", "kopia:", "ämne:", "datum:",
        };

        /// <summary>
        /// The part of the email worth judging the language by: header lines of quoted
        /// emails ("From:", "To:" - in the language of Outlook, not of the email) and
        /// lines with addresses or links are dropped, and only the first 600 characters
        /// of what remains are used. What is nearest the top - what Olle has written and
        /// the start of the email he replies to - decides, not the long English
        /// disclaimers further down (found on 10 November 2026: a Swedish reply came out
        /// undecided because of them).
        /// </summary>
        public static string LanguageSample(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new StringBuilder();
            foreach (string raw in text.Split(new[] { "\r\n", "\n", "\r", "\v" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string line = raw.Trim();
                if (line.Length == 0) continue;
                string lower = line.ToLowerInvariant();
                if (HeaderLabels.Any(h => lower.StartsWith(h))) continue;
                if (lower.Contains("@") || lower.Contains("http")) continue;
                sb.Append(line).Append(' ');
                if (sb.Length >= 600) break;
            }
            return sb.Length > 600 ? sb.ToString(0, 600) : sb.ToString();
        }

        /// <summary>
        /// Swedish or English from the email's text (including the quoted thread), or
        /// null when there is too little to tell - the caller then uses the last choice.
        /// </summary>
        public static AvailabilityLanguage? DetectLanguage(string text)
        {
            int sv, en;
            return DetectLanguage(text, out sv, out en);
        }

        /// <summary>The same, also giving the Swedish and English word counts, for the log.</summary>
        public static AvailabilityLanguage? DetectLanguage(string text, out int sv, out int en)
        {
            sv = 0;
            en = 0;
            text = LanguageSample(text);
            if (string.IsNullOrWhiteSpace(text)) return null;
            var word = new StringBuilder();
            foreach (char c in text + " ")
            {
                if (char.IsLetter(c)) { word.Append(c); continue; }
                if (word.Length > 0)
                {
                    string w = word.ToString();
                    if (SvWords.Contains(w)) sv++;
                    else if (EnWords.Contains(w)) en++;
                    if (w.IndexOfAny(new[] { 'å', 'ä', 'ö', 'Å', 'Ä', 'Ö' }) >= 0) sv++;
                    word.Clear();

                    // Read from the top, like a person would: once one language clearly
                    // leads, that is the email's language, whatever disclaimer follows.
                    if (sv - en >= DecisiveLead) return AvailabilityLanguage.Swedish;
                    if (en - sv >= DecisiveLead) return AvailabilityLanguage.English;
                }
            }
            if (sv + en < 3) return null;
            if (sv > en * 1.2) return AvailabilityLanguage.Swedish;
            if (en > sv * 1.2) return AvailabilityLanguage.English;
            return null;
        }
    }

    internal struct TimeRange
    {
        public TimeRange(DateTime start, DateTime end)
        {
            Start = start;
            End = end;
        }

        public DateTime Start { get; }
        public DateTime End { get; }

        public override string ToString() => Start.ToString("ddd dd HH:mm") + "-" + End.ToString("HH:mm");
    }

    internal sealed class AvailabilityRules
    {
        public TimeSpan WorkStart { get; set; } = new TimeSpan(8, 0, 0);
        public TimeSpan WorkEnd { get; set; } = new TimeSpan(17, 0, 0);
        public HashSet<DayOfWeek> WorkDays { get; set; } = new HashSet<DayOfWeek>
        {
            DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday,
        };
        public TimeSpan Buffer { get; set; } = TimeSpan.Zero;
        /// <summary>Lunch, never offered, on every working day. Null start or end means no lunch break.</summary>
        public TimeSpan? LunchStart { get; set; }
        public TimeSpan? LunchEnd { get; set; }
    }

    internal enum AvailabilityPeriod { ThisWeek, NextWeek, NextTwoWeeks }

    internal enum AvailabilityFormat { List, Sentence }

    internal enum AvailabilityLanguage { Swedish, English }
}
