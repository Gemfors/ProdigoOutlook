using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ProdigoOutlook.Features.PreSendChecks
{
    /// <summary>
    /// The pre-send checks without Outlook: the email's text, attachments and
    /// recipients in, findings out. Tested in test\LogicTests. PreSendChecksOnSend reads
    /// the email and shows the warning; PreSendRulesFile keeps the rules.
    ///
    /// Agreed with Olle on 10-11 November 2026:
    /// - Missing attachment, in Swedish and English. Prodigo takes this over from
    ///   Outlook, whose own reminder does not understand Swedish; Olle switches
    ///   Outlook's off once this works.
    /// - A reply to more than 10 people. Outlook does not record whether an email was a
    ///   "Reply all", so a reply that goes to many is what counts.
    /// - Recipients both inside HiQ (hiq.se, hiq.fi, hiq.de, zoolengineering.se) and
    ///   outside, naming the outside domains.
    /// - A weekday that does not match its date ("tisdag 12 nov" on a Thursday).
    /// - Emails only, not meeting invites. Empty subject is left to Outlook's own
    ///   question. No account check: Olle sends from one account.
    /// Only Olle's own text counts for the text checks, not the quoted email below it.
    /// </summary>
    internal static class PreSendCheck
    {
        public static List<PreSendFinding> Run(PreSendInput input, PreSendRules rules)
        {
            var findings = new List<PreSendFinding>();
            string own = OwnText(input.Body);

            if (rules.CheckAttachment && input.RealAttachments == 0)
            {
                string word = FirstAttachmentWord(own, rules.AttachmentWords);
                if (word != null)
                    findings.Add(new PreSendFinding(PreSendKind.Attachment,
                        "You mention an attachment (\"" + word + "\"), but nothing is attached."));
            }

            List<string> addresses = (input.Recipients ?? new List<string>())
                .Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => a.Trim().ToLowerInvariant()).Distinct().ToList();

            if (rules.CheckReplyAll && input.IsReply && addresses.Count > rules.ReplyAllLimit)
                findings.Add(new PreSendFinding(PreSendKind.ReplyAll,
                    "This reply goes to " + addresses.Count + " people."));

            if (rules.CheckMixed)
            {
                List<string> domains = addresses.Select(DomainOf).Where(d => d != null).ToList();
                List<string> outside = domains.Where(d => !IsInternal(d, rules.InternalDomains)).Distinct().ToList();
                bool anyInside = domains.Any(d => IsInternal(d, rules.InternalDomains));
                if (anyInside && outside.Count > 0)
                    findings.Add(new PreSendFinding(PreSendKind.Mixed,
                        "This goes to people both inside HiQ and outside: " +
                        string.Join(", ", outside.Take(5)) + (outside.Count > 5 ? " and " + (outside.Count - 5) + " more" : "") + "."));
            }

            if (rules.CheckDates)
                foreach (string problem in DateProblems(own, input.Today))
                    findings.Add(new PreSendFinding(PreSendKind.Date, problem));

            return findings;
        }

        // ---------------------------------------------------------------
        // Own text
        // ---------------------------------------------------------------

        private static readonly Regex QuoteStart = new Regex(
            @"^\s*(-{2,}\s*(original message|ursprungligt meddelande|forwarded message|vidarebefordrat meddelande)\s*-{2,}" +
            @"|(on|den)\s.{5,120}\s(wrote|skrev):\s*$)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex HeaderFrom = new Regex(@"^\s*(from|från)\s*:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex HeaderNext = new Regex(@"^\s*(sent|skickat|date|datum|to|till)\s*:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// The part Olle wrote: everything above the first quoted email - a "From:" /
        /// "Från:" line followed by "Sent:" or similar, an "Original message" divider,
        /// or "On ... wrote:".
        /// </summary>
        public static string OwnText(string body)
        {
            if (string.IsNullOrEmpty(body)) return string.Empty;
            string[] lines = body.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (QuoteStart.IsMatch(lines[i])) return string.Join("\n", lines.Take(i));
                if (HeaderFrom.IsMatch(lines[i]))
                    for (int j = i + 1; j < Math.Min(lines.Length, i + 4); j++)
                        if (HeaderNext.IsMatch(lines[j])) return string.Join("\n", lines.Take(i));
            }
            return body;
        }

        // ---------------------------------------------------------------
        // Attachment
        // ---------------------------------------------------------------

        /// <summary>
        /// The first word in the text matching one of the attachment words, or null. An
        /// entry ending in * matches any word starting with it ("bifoga*" matches
        /// bifogar, bifogad, bifogade); otherwise the whole word must match.
        /// </summary>
        public static string FirstAttachmentWord(string text, IEnumerable<string> words)
        {
            if (string.IsNullOrEmpty(text) || words == null) return null;
            var entries = words.Select(w => (w ?? string.Empty).Trim().ToLowerInvariant()).Where(w => w.Length > 0).ToList();
            foreach (Match m in Regex.Matches(text, @"\p{L}+"))
            {
                string w = m.Value.ToLowerInvariant();
                foreach (string e in entries)
                {
                    bool hit = e.EndsWith("*") ? w.StartsWith(e.TrimEnd('*')) : w == e;
                    if (hit) return m.Value;
                }
            }
            return null;
        }

        // ---------------------------------------------------------------
        // Recipients
        // ---------------------------------------------------------------

        private static readonly Regex ReplyPrefix = new Regex(@"^\s*(re|sv|aw|antw|vs)\s*:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// A reply, judged by its subject ("RE:", "SV:", "AW:"...): Outlook keeps no
        /// record of whether Reply or Reply all was pressed. Forwards ("FW:", "VB:")
        /// are not replies - they go to whoever they are forwarded to.
        /// </summary>
        public static bool IsReplySubject(string subject) => !string.IsNullOrEmpty(subject) && ReplyPrefix.IsMatch(subject);

        public static string DomainOf(string address)
        {
            if (string.IsNullOrEmpty(address)) return null;
            int at = address.LastIndexOf('@');
            return at < 0 || at == address.Length - 1 ? null : address.Substring(at + 1).Trim().ToLowerInvariant();
        }

        /// <summary>The domain itself or any subdomain of it ("se.hiq.se" counts as hiq.se).</summary>
        public static bool IsInternal(string domain, IEnumerable<string> internalDomains) =>
            internalDomains != null && internalDomains
                .Select(d => (d ?? string.Empty).Trim().TrimStart('@').ToLowerInvariant())
                .Where(d => d.Length > 0)
                .Any(d => domain == d || domain.EndsWith("." + d));

        // ---------------------------------------------------------------
        // Weekday and date
        // ---------------------------------------------------------------

        private static readonly Dictionary<string, DayOfWeek> Weekdays = new Dictionary<string, DayOfWeek>(StringComparer.OrdinalIgnoreCase)
        {
            { "måndag", DayOfWeek.Monday }, { "mån", DayOfWeek.Monday },
            { "tisdag", DayOfWeek.Tuesday }, { "tis", DayOfWeek.Tuesday },
            { "onsdag", DayOfWeek.Wednesday }, { "ons", DayOfWeek.Wednesday },
            { "torsdag", DayOfWeek.Thursday }, { "tors", DayOfWeek.Thursday }, { "tor", DayOfWeek.Thursday },
            { "fredag", DayOfWeek.Friday }, { "fre", DayOfWeek.Friday },
            { "lördag", DayOfWeek.Saturday }, { "lör", DayOfWeek.Saturday },
            { "söndag", DayOfWeek.Sunday }, { "sön", DayOfWeek.Sunday },
            { "monday", DayOfWeek.Monday }, { "mon", DayOfWeek.Monday },
            { "tuesday", DayOfWeek.Tuesday }, { "tue", DayOfWeek.Tuesday }, { "tues", DayOfWeek.Tuesday },
            { "wednesday", DayOfWeek.Wednesday }, { "wed", DayOfWeek.Wednesday },
            { "thursday", DayOfWeek.Thursday }, { "thu", DayOfWeek.Thursday }, { "thur", DayOfWeek.Thursday }, { "thurs", DayOfWeek.Thursday },
            { "friday", DayOfWeek.Friday }, { "fri", DayOfWeek.Friday },
            { "saturday", DayOfWeek.Saturday }, { "sat", DayOfWeek.Saturday },
            { "sunday", DayOfWeek.Sunday }, { "sun", DayOfWeek.Sunday },
        };

        private static readonly Dictionary<string, int> Months = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "januari", 1 }, { "january", 1 }, { "jan", 1 },
            { "februari", 2 }, { "february", 2 }, { "feb", 2 },
            { "mars", 3 }, { "march", 3 }, { "mar", 3 },
            { "april", 4 }, { "apr", 4 },
            { "maj", 5 }, { "may", 5 },
            { "juni", 6 }, { "june", 6 }, { "jun", 6 },
            { "juli", 7 }, { "july", 7 }, { "jul", 7 },
            { "augusti", 8 }, { "august", 8 }, { "aug", 8 },
            { "september", 9 }, { "sept", 9 }, { "sep", 9 },
            { "oktober", 10 }, { "october", 10 }, { "okt", 10 }, { "oct", 10 },
            { "november", 11 }, { "nov", 11 },
            { "december", 12 }, { "dec", 12 },
        };

        private static readonly string WeekdayAlt = string.Join("|", Weekdays.Keys.OrderByDescending(k => k.Length));
        private static readonly string MonthAlt = string.Join("|", Months.Keys.OrderByDescending(k => k.Length));

        // "tisdag 12 nov", "tisdagen den 12:e november", "tis 12/11", "Tuesday 12 November",
        // "Tue, 12th of Nov", each optionally followed by a year.
        private static readonly Regex DayFirst = new Regex(
            @"\b(?<wd>" + WeekdayAlt + @")(?:en)?\.?,?\s+(?:den\s+|the\s+)?(?<d>\d{1,2})(?::?e|:?a|st|nd|rd|th)?" +
            @"(?:\s+(?:of\s+)?(?<m>" + MonthAlt + @")\.?|/(?<mn>\d{1,2})(?!/\d))(?:\s+(?<y>\d{4}))?\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // "Tuesday, November 12", "Tue Nov 12th, 2026".
        private static readonly Regex MonthFirst = new Regex(
            @"\b(?<wd>" + WeekdayAlt + @")\.?,?\s+(?<m>" + MonthAlt + @")\.?\s+(?<d>\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(?<y>\d{4}))?\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly string[] EnDayNames = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

        /// <summary>
        /// One sentence per weekday-and-date in the text that do not agree. Without a
        /// year, the date is taken as the one nearest to today (so "fre 2 jan" written
        /// in December means next January).
        /// </summary>
        public static List<string> DateProblems(string text, DateTime today)
        {
            var problems = new List<string>();
            if (string.IsNullOrEmpty(text)) return problems;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Regex rx in new[] { DayFirst, MonthFirst })
            {
                foreach (Match m in rx.Matches(text))
                {
                    DayOfWeek written = Weekdays[m.Groups["wd"].Value];
                    int day = int.Parse(m.Groups["d"].Value, CultureInfo.InvariantCulture);
                    int month = m.Groups["m"].Success ? Months[m.Groups["m"].Value]
                              : int.Parse(m.Groups["mn"].Value, CultureInfo.InvariantCulture);
                    int? year = m.Groups["y"].Success ? int.Parse(m.Groups["y"].Value, CultureInfo.InvariantCulture) : (int?)null;

                    DateTime? date = Resolve(day, month, year, today);
                    if (!date.HasValue || date.Value.DayOfWeek == written) continue;
                    if (!seen.Add(m.Value)) continue;

                    problems.Add("\"" + m.Value.Trim() + "\": " + date.Value.Day + " " +
                                 CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(date.Value.Month) + " " +
                                 date.Value.Year + " is a " + EnDayNames[(int)date.Value.DayOfWeek] + ".");
                }
            }
            return problems;
        }

        private static DateTime? Resolve(int day, int month, int? year, DateTime today)
        {
            if (month < 1 || month > 12 || day < 1) return null;
            IEnumerable<int> years = year.HasValue ? new[] { year.Value } : new[] { today.Year - 1, today.Year, today.Year + 1 };
            DateTime? best = null;
            foreach (int y in years)
            {
                if (y < 1900 || y > 2200 || day > DateTime.DaysInMonth(y, month)) continue;
                var d = new DateTime(y, month, day);
                if (!best.HasValue || Math.Abs((d - today.Date).TotalDays) < Math.Abs((best.Value - today.Date).TotalDays))
                    best = d;
            }
            return best;
        }
    }

    internal sealed class PreSendRules
    {
        public bool CheckAttachment { get; set; } = true;
        public bool CheckReplyAll { get; set; } = true;
        public bool CheckMixed { get; set; } = true;
        public bool CheckDates { get; set; } = true;
        public List<string> InternalDomains { get; set; } = new List<string> { "hiq.se", "hiq.fi", "hiq.de", "zoolengineering.se" };
        public int ReplyAllLimit { get; set; } = 10;
        public List<string> AttachmentWords { get; set; } = new List<string>
        {
            "bifoga*", "bilaga", "bilagan", "bilagor", "bilagorna", "bilagd", "bilagda",
            "attach*", "enclosed", "enclosing",
        };
    }

    internal sealed class PreSendInput
    {
        /// <summary>The whole body as plain text; the quoted part is cut off here.</summary>
        public string Body { get; set; }
        /// <summary>Attachments the reader would see - not signature images or other pictures inside the text.</summary>
        public int RealAttachments { get; set; }
        /// <summary>Recipients' email addresses (To, Cc and Bcc).</summary>
        public List<string> Recipients { get; set; }
        public bool IsReply { get; set; }
        public DateTime Today { get; set; }
    }

    internal enum PreSendKind { Attachment, ReplyAll, Mixed, Date }

    internal sealed class PreSendFinding
    {
        public PreSendFinding(PreSendKind kind, string message)
        {
            Kind = kind;
            Message = message;
        }

        public PreSendKind Kind { get; }
        public string Message { get; }
    }
}
