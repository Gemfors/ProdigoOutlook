using System;
using System.Collections.Generic;
using ProdigoOutlook.Core;
using Office = Microsoft.Office.Core;

namespace ProdigoOutlook.Features.Availability
{
    /// <summary>
    /// PHASE 3 - Insert availability. In an email being written (its own window, or a
    /// reply in the reading pane), the button reads Olle's own calendar, shows the
    /// free times in a small window with a live preview, and inserts them at the
    /// cursor as a list or a sentence, in Swedish or English.
    ///
    /// It only reads the calendar and inserts text: it never changes the calendar and
    /// never sends anything. The rules were agreed with Olle on 10 November 2026; see
    /// AvailabilityText.
    /// </summary>
    internal static class Availability
    {
        public static void Insert(Office.IRibbonControl control)
        {
            object window = Ol.ContextWindow(control);
            object item, editor;
            if (!AvailabilityCalendar.TryGetDraft(window, out item, out editor))
            {
                Ol.Info("Click in an email you're writing first - a new email, a reply or a forward - " +
                        "then use Insert Availability.", "Insert Availability");
                return;
            }

            OutlookSettings settings = OutlookSettings.Current;
            AvailabilityRules rules = AvailabilityText.Rules(settings.AvailabilityWorkStart, settings.AvailabilityWorkEnd,
                settings.AvailabilityWorkDays, settings.AvailabilityBuffer,
                settings.AvailabilityLunchStart, settings.AvailabilityLunchEnd);

            DateTime now = DateTime.Now;
            List<TimeRange> busy = AvailabilityCalendar.Busy(now.Date, now.Date.AddDays(15));

            int sv, en;
            AvailabilityLanguage? detected = AvailabilityText.DetectLanguage(AvailabilityCalendar.BodyText(item, editor), out sv, out en);
            AvailabilityLanguage language = detected
                ?? (settings.AvailabilityLastLanguage == "en" ? AvailabilityLanguage.English : AvailabilityLanguage.Swedish);
            Log.Info("Insert availability: language " + (detected.HasValue ? "detected " + detected.Value : "undecided, using last choice " + language) +
                     " (Swedish words " + sv + ", English words " + en + ")");

            AvailabilityChoice choice = AvailabilityForm.Ask(Ol.OwnerWindow(window), busy, rules, now, language, settings);
            if (choice == null) return;

            // A formatted email gets real bullets; a plain-text one keeps the "• " lines
            // the preview showed.
            bool realBullets = choice.Format == AvailabilityFormat.List && !AvailabilityCalendar.IsPlainText(item);
            if (realBullets)
                AvailabilityCalendar.Insert(editor, choice.TextWithoutBullets, choice.HasOpening ? 1 : 0);
            else
                AvailabilityCalendar.Insert(editor, choice.Text, -1);
            Log.Info("Insert availability: inserted " + choice.RangeCount + " time range(s)");

            OutlookSettings updated = settings.Copy();
            updated.AvailabilityLastLength = choice.LengthMinutes;
            updated.AvailabilityLastPeriod = PeriodKey(choice.Period);
            updated.AvailabilityLastFormat = choice.Format == AvailabilityFormat.Sentence ? "sentence" : "list";
            updated.AvailabilityLastLanguage = choice.Language == AvailabilityLanguage.English ? "en" : "sv";
            updated.SaveAsCurrent();
        }

        public static string PeriodKey(AvailabilityPeriod p) =>
            p == AvailabilityPeriod.NextWeek ? "nextweek" : p == AvailabilityPeriod.NextTwoWeeks ? "next2weeks" : "thisweek";

        public static AvailabilityPeriod PeriodFromKey(string key) =>
            key == "nextweek" ? AvailabilityPeriod.NextWeek : key == "next2weeks" ? AvailabilityPeriod.NextTwoWeeks : AvailabilityPeriod.ThisWeek;
    }

    internal sealed class AvailabilityChoice
    {
        /// <summary>As previewed: list lines start with "• ".</summary>
        public string Text { get; set; }
        /// <summary>The same with no bullet characters, for real Word bullets.</summary>
        public string TextWithoutBullets { get; set; }
        public bool HasOpening { get; set; }
        public int RangeCount { get; set; }
        public int LengthMinutes { get; set; }
        public AvailabilityPeriod Period { get; set; }
        public AvailabilityFormat Format { get; set; }
        public AvailabilityLanguage Language { get; set; }
    }
}
