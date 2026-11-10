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
                settings.AvailabilityWorkDays, settings.AvailabilityBuffer);

            DateTime now = DateTime.Now;
            List<TimeRange> busy = AvailabilityCalendar.Busy(now.Date, now.Date.AddDays(15));

            AvailabilityLanguage language = AvailabilityText.DetectLanguage(AvailabilityCalendar.BodyText(item))
                ?? (settings.AvailabilityLastLanguage == "en" ? AvailabilityLanguage.English : AvailabilityLanguage.Swedish);

            AvailabilityChoice choice = AvailabilityForm.Ask(Ol.OwnerWindow(window), busy, rules, now, language, settings);
            if (choice == null) return;

            AvailabilityCalendar.Insert(editor, choice.Text);
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
        public string Text { get; set; }
        public int RangeCount { get; set; }
        public int LengthMinutes { get; set; }
        public AvailabilityPeriod Period { get; set; }
        public AvailabilityFormat Format { get; set; }
        public AvailabilityLanguage Language { get; set; }
    }
}
