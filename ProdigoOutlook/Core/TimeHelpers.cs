using System;
using System.Globalization;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Date and time odds and ends shared by the features. Pure logic, no Outlook
    /// types, so test\LogicTests compiles it as it is.
    /// </summary>
    internal static class TimeHelpers
    {
        /// <summary>
        /// Outlook's way of saying "no date": properties such as DeferredDeliveryTime
        /// read as 1 January 4501 when they are not set, rather than as null.
        /// </summary>
        public static readonly DateTime OutlookNoDate = new DateTime(4501, 1, 1);

        public static bool IsNoDate(DateTime value) => value.Year >= 4500 || value == DateTime.MinValue;

        /// <summary>The English the UI is written in, regardless of Windows' language.</summary>
        public static readonly CultureInfo UiCulture = CultureInfo.GetCultureInfo("en-GB");

        /// <summary>"Thu 5 Nov 08:00" - short, unambiguous, and the same everywhere in the UI.</summary>
        public static string Friendly(DateTime local) => local.ToString("ddd d MMM HH:mm", UiCulture);

        /// <summary>"Thu 5 Nov" without the time.</summary>
        public static string FriendlyDate(DateTime local) => local.ToString("ddd d MMM", UiCulture);

        /// <summary>
        /// Round-trip text for a moment in time, always in UTC, so a value written on a
        /// laptop in Stockholm reads back correctly on a laptop in Helsinki.
        /// </summary>
        public static string ToStored(DateTime local) =>
            local.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

        /// <summary>Parses <see cref="ToStored"/> back to local time, or null.</summary>
        public static DateTime? FromStored(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            DateTime utc;
            if (DateTime.TryParseExact(text.Trim(), "yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out utc))
                return DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToLocalTime();
            return null;
        }

        /// <summary>Drops seconds and below. Delivery times are chosen to the minute.</summary>
        public static DateTime TruncateToMinute(DateTime t) =>
            new DateTime(t.Year, t.Month, t.Day, t.Hour, t.Minute, 0, t.Kind);

        public static bool IsWeekend(DateTime d) =>
            d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday;

        /// <summary>True when two times are within a minute of each other.</summary>
        public static bool SameMinute(DateTime a, DateTime b) => Math.Abs((a - b).TotalSeconds) < 60;
    }
}
