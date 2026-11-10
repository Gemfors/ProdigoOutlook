namespace ProdigoOutlook.Core
{
    /// <summary>
    /// What a fresh install starts from, for everything the user can change in
    /// Settings. The same pattern as ProdigoExcel's Settings.cs: what is in force lives
    /// in the settings file (OutlookSettings); these are only the starting values.
    /// A version for someone else changes these and nothing else.
    /// </summary>
    internal static class Defaults
    {
        /// <summary>
        /// The Send later preset buttons, top to bottom. One per line in the settings
        /// file as "sendlater.preset = ...". Grammar in SendLaterTime.ParsePreset:
        /// "+1h", "+30m", "today 13:00", "tomorrow 08:00", "nextworkday 08:00",
        /// or a weekday such as "monday 08:00".
        /// </summary>
        public static readonly string[] SendLaterPresets =
        {
            "+1h",
            "tomorrow 08:00",
            "nextworkday 08:00",
            "monday 08:00",
        };

        /// <summary>Warn when a delayed meeting includes a meeting room.</summary>
        public const bool SendLaterRoomWarning = true;

        // Insert availability (phase 3), agreed with Olle on 10 November 2026.
        public const string AvailabilityWorkStart = "08:00";
        public const string AvailabilityWorkEnd = "17:00";
        public const string AvailabilityWorkDays = "mon,tue,wed,thu,fri";
        /// <summary>Minutes kept clear before and after each meeting.</summary>
        public const int AvailabilityBuffer = 0;
        public const string AvailabilityOpeningListEn = "I'm available at the following times:";
        public const string AvailabilityOpeningListSv = "Jag är ledig följande tider:";
        public const string AvailabilityOpeningSentenceEn = "I'm free";
        public const string AvailabilityOpeningSentenceSv = "Jag kan";
    }
}
