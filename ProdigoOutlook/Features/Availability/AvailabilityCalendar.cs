using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using ProdigoOutlook.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Features.Availability
{
    /// <summary>
    /// The Outlook side of Insert availability: reads busy times from Olle's own
    /// default calendar, and puts text into the email being written. Reading only -
    /// nothing here changes the calendar or sends anything.
    ///
    /// Calendar items are not Outbox items, so reading them through Items is safe
    /// (the Outbox trap in CLAUDE.md does not apply here). Recurring meetings are
    /// expanded with IncludeRecurrences, which needs Sort("[Start]") first and
    /// GetFirst/GetNext rather than Count. The Restrict date strings use the current
    /// culture's short format, which is what Outlook expects; checked against Olle's
    /// real calendar on 10 November 2026.
    /// </summary>
    internal static class AvailabilityCalendar
    {
        /// <summary>Everything not marked Free between the two times, in time order.</summary>
        public static List<TimeRange> Busy(DateTime from, DateTime to)
        {
            var result = new List<TimeRange>();
            using (var com = new ComScope())
            {
                Outlook.NameSpace ns = com.Track(Ol.App.Session);
                Outlook.MAPIFolder calendar = com.Track(ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar));
                Outlook.Items items = com.Track(calendar.Items);
                items.Sort("[Start]");
                items.IncludeRecurrences = true;

                string filter = "[Start] < '" + to.ToString("g", CultureInfo.CurrentCulture) +
                                "' AND [End] > '" + from.ToString("g", CultureInfo.CurrentCulture) + "'";
                Outlook.Items found = com.Track(items.Restrict(filter));

                object item = found.GetFirst();
                int guard = 0;
                while (item != null && guard++ < 5000)
                {
                    try
                    {
                        if (item is Outlook.AppointmentItem appt && appt.BusyStatus != Outlook.OlBusyStatus.olFree)
                            result.Add(new TimeRange(appt.Start, appt.End));
                    }
                    catch (Exception ex)
                    {
                        Log.Warn("Insert availability: skipped a calendar item", ex);
                    }
                    finally
                    {
                        ComScope.Release(item);
                    }
                    item = found.GetNext();
                }
            }
            return result;
        }

        /// <summary>
        /// The email being written in this window, if there is one: an open compose
        /// window, or a reply written in the reading pane of the main window.
        /// Neither item nor editor is released by the caller's side; see Insert.
        /// </summary>
        public static bool TryGetDraft(object window, out object item, out object wordEditor)
        {
            item = null;
            wordEditor = null;
            try
            {
                if (window is Outlook.Inspector inspector)
                {
                    if (!(inspector.CurrentItem is Outlook.MailItem mail) || mail.Sent) return false;
                    item = mail;
                    wordEditor = inspector.WordEditor;
                }
                else if (window is Outlook.Explorer explorer)
                {
                    item = explorer.ActiveInlineResponse;
                    if (!(item is Outlook.MailItem)) { item = null; return false; }
                    wordEditor = explorer.ActiveInlineResponseWordEditor;
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Insert availability: could not reach the email being written", ex);
                return false;
            }
            return item != null && wordEditor != null;
        }

        /// <summary>
        /// The email's text as it is on screen, read from its editor - the item's Body
        /// can lag behind while the email is being written. Falls back to Body.
        /// </summary>
        public static string BodyText(object item, object wordEditor)
        {
            dynamic content = null;
            try
            {
                dynamic doc = wordEditor;
                content = doc.Content;
                string text = content.Text as string;
                if (!string.IsNullOrWhiteSpace(text)) return text;
            }
            catch (Exception ex)
            {
                Log.Warn("Insert availability: could not read the email's text from its editor", ex);
            }
            finally
            {
                if (content != null && Marshal.IsComObject((object)content)) Marshal.ReleaseComObject((object)content);
            }
            try { return (item as Outlook.MailItem)?.Body ?? string.Empty; }
            catch { return string.Empty; }
        }

        /// <summary>True for a plain-text email, where real bullets are not possible.</summary>
        public static bool IsPlainText(object item)
        {
            try { return (item as Outlook.MailItem)?.BodyFormat == Outlook.OlBodyFormat.olFormatPlain; }
            catch { return false; }
        }

        /// <summary>
        /// Inserts the text at the cursor, replacing any selected text, and leaves the
        /// cursor after it. "\n" becomes a new paragraph. From line
        /// <paramref name="firstBulletLine"/> on (0-based), the paragraphs get real Word
        /// bullets, as with Outlook's own Bullets button; -1 means none. If the bullets
        /// cannot be applied, those lines get a "• " instead, so the list never comes
        /// out bare. Goes through Word's own editor, late-bound so no Word reference is
        /// needed.
        /// </summary>
        public static void Insert(object wordEditor, string text, int firstBulletLine)
        {
            dynamic doc = wordEditor;
            dynamic windows = null, window = null, selection = null, range = null;
            try
            {
                windows = doc.Windows;
                window = windows[1];
                selection = window.Selection;
                range = selection.Range;
                range.Text = text.Replace("\r\n", "\n").Replace('\n', '\r');
                int end = range.End;

                if (firstBulletLine >= 0)
                {
                    try
                    {
                        ApplyBullets(doc, range, firstBulletLine);
                    }
                    catch (Exception ex)
                    {
                        Log.Warn("Insert availability: real bullets could not be applied; using a bullet character", ex);
                        PrefixBullets(range, firstBulletLine);
                    }
                    end = range.End;
                }

                selection.SetRange(end, end);
            }
            finally
            {
                foreach (object o in new object[] { range, selection, window, windows })
                    if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
            }
        }

        private const int WdBulletGallery = 1;
        private const int WdListApplyToWholeList = 0;

        private static void ApplyBullets(dynamic doc, dynamic range, int firstBulletLine)
        {
            dynamic paragraphs = range.Paragraphs;
            dynamic first = paragraphs[firstBulletLine + 1];
            dynamic listRange = doc.Range(first.Range.Start, range.End);
            dynamic galleries = doc.Application.ListGalleries;
            dynamic template = galleries[WdBulletGallery].ListTemplates[1];
            // ApplyListTemplate rather than ApplyBulletDefault: the latter toggles, so it
            // would remove bullets if the cursor was already in a bulleted list.
            listRange.ListFormat.ApplyListTemplate(template, false, WdListApplyToWholeList);
        }

        private static void PrefixBullets(dynamic range, int firstBulletLine)
        {
            dynamic paragraphs = range.Paragraphs;
            int count = paragraphs.Count;
            for (int i = firstBulletLine + 1; i <= count; i++)
                paragraphs[i].Range.InsertBefore("• ");
        }
    }
}
