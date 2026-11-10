using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ProdigoOutlook.Features.Availability;
using ProdigoOutlook.Features.SendLater;
using Office = Microsoft.Office.Core;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Prodigo > Settings. Like ProdigoExcel's settings dialog it edits a copy and
    /// applies it only on OK, so Cancel really cancels.
    ///
    /// Features: one checkbox each. Planned features are listed greyed out, so the
    /// roadmap is visible from inside Outlook. Switching a built feature off hides its
    /// buttons and takes it out of sending straight away - the escape hatch if a build
    /// misbehaves.
    /// </summary>
    internal sealed class SettingsForm : Form
    {
        private readonly OutlookSettings _edit = OutlookSettings.Current.Copy();
        private readonly Dictionary<string, CheckBox> _featureBoxes = new Dictionary<string, CheckBox>();
        private TextBox _presets;
        private CheckBox _roomWarning;
        private TextBox _workStart, _workEnd, _openListEn, _openListSv, _openSentEn, _openSentSv;
        private readonly Dictionary<DayOfWeek, CheckBox> _dayBoxes = new Dictionary<DayOfWeek, CheckBox>();
        private NumericUpDown _buffer;
        private Label _error;

        public static void Open(Office.IRibbonControl control)
        {
            using (var form = new SettingsForm())
            {
                if (form.ShowDialog(Ol.OwnerWindow(Ol.ContextWindow(control))) == DialogResult.OK)
                    Globals.ThisAddIn.RibbonInstance?.InvalidateAll();
            }
        }

        private SettingsForm()
        {
            FormStyle.ApplyDialog(this, "Prodigo Settings");
            Build();
        }

        private void Build()
        {
            const int left = 16, width = 460;
            int top = 12;

            var features = new GroupBox { Text = "Features", Left = left, Top = top, Width = width };
            int y = 22;
            foreach (FeatureInfo f in Feature.All)
            {
                var box = new CheckBox
                {
                    Text = f.Name + (f.Built ? string.Empty : "   (planned, phase " + f.Phase + ")"),
                    Left = 12, Top = y, Width = width - 24, Height = 22,
                    Checked = f.Built && _edit.IsFeatureOn(f.Id),
                    Enabled = f.Built,
                };
                new ToolTip().SetToolTip(box, f.Description);
                _featureBoxes[f.Id] = box;
                features.Controls.Add(box);
                y += 24;
            }
            features.Height = y + 8;
            Controls.Add(features);
            top += features.Height + 10;

            var sendLater = new GroupBox { Text = "Send later", Left = left, Top = top, Width = width };
            sendLater.Controls.Add(new Label
            {
                Text = "Preset buttons, one per line, top to bottom. For example: +1h, +30m, today 13:00, " +
                       "tomorrow 08:00, nextworkday 08:00, monday 08:00.",
                Left = 12, Top = 22, Width = width - 24, Height = 34,
            });
            _presets = new TextBox
            {
                Multiline = true, ScrollBars = ScrollBars.Vertical, AcceptsReturn = true,
                Left = 12, Top = 60, Width = width - 24, Height = 96,
                Text = string.Join(Environment.NewLine, _edit.SendLaterPresets),
            };
            sendLater.Controls.Add(_presets);
            _roomWarning = new CheckBox
            {
                Text = "Warn when a delayed meeting includes a meeting room",
                Left = 12, Top = 162, Width = width - 24, Height = 22,
                Checked = _edit.SendLaterRoomWarning,
            };
            sendLater.Controls.Add(_roomWarning);
            sendLater.Height = 194;
            Controls.Add(sendLater);
            top += sendLater.Height + 10;

            // Insert availability, in a second column so the window stays short enough
            // for a laptop screen at 150% scaling.
            int col2 = left * 2 + width;
            var avail = new GroupBox { Text = "Insert availability", Left = col2, Top = 12, Width = width };
            int ay = 24;
            avail.Controls.Add(new Label { Text = "Working hours:", Left = 12, Top = ay + 4, Width = 110 });
            _workStart = new TextBox { Left = 126, Top = ay, Width = 60, Text = _edit.AvailabilityWorkStart };
            avail.Controls.Add(_workStart);
            avail.Controls.Add(new Label { Text = "to", Left = 192, Top = ay + 4, Width = 24 });
            _workEnd = new TextBox { Left = 218, Top = ay, Width = 60, Text = _edit.AvailabilityWorkEnd };
            avail.Controls.Add(_workEnd);
            ay += 32;

            avail.Controls.Add(new Label { Text = "Working days:", Left = 12, Top = ay + 4, Width = 110 });
            HashSet<DayOfWeek> days = AvailabilityText.ParseDays(_edit.AvailabilityWorkDays)
                                      ?? AvailabilityText.ParseDays(Defaults.AvailabilityWorkDays);
            DayOfWeek[] order = { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday,
                                  DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday };
            string[] dayLabels = { "Mo", "Tu", "We", "Th", "Fr", "Sa", "Su" };
            for (int i = 0; i < order.Length; i++)
            {
                var box = new CheckBox { Text = dayLabels[i], Left = 126 + i * 44, Top = ay + 2, Width = 44, Checked = days.Contains(order[i]) };
                _dayBoxes[order[i]] = box;
                avail.Controls.Add(box);
            }
            ay += 32;

            avail.Controls.Add(new Label { Text = "Keep clear before and after meetings (minutes):", Left = 12, Top = ay + 4, Width = 290 });
            _buffer = new NumericUpDown { Left = 306, Top = ay, Width = 60, Minimum = 0, Maximum = 120, Increment = 5, Value = Math.Min(120, _edit.AvailabilityBuffer) };
            avail.Controls.Add(_buffer);
            ay += 36;

            avail.Controls.Add(new Label { Text = "Opening lines (leave empty for none):", Left = 12, Top = ay, Width = width - 24 });
            ay += 22;
            _openListEn = OpeningBox(avail, "List, English:", _edit.AvailabilityOpeningListEn, ref ay);
            _openListSv = OpeningBox(avail, "List, Swedish:", _edit.AvailabilityOpeningListSv, ref ay);
            _openSentEn = OpeningBox(avail, "Sentence, English:", _edit.AvailabilityOpeningSentenceEn, ref ay);
            _openSentSv = OpeningBox(avail, "Sentence, Swedish:", _edit.AvailabilityOpeningSentenceSv, ref ay);
            avail.Height = ay + 8;
            Controls.Add(avail);
            top = Math.Max(top, 12 + avail.Height + 10);

            int total = left * 3 + width * 2;
            _error = new Label { Left = left, Top = top, Width = total - left * 2, Height = 34, ForeColor = Color.Firebrick };
            Controls.Add(_error);
            top += 38;

            var openLog = new Button { Text = "Open &log folder", Left = left, Top = top, Width = 130, Height = 28 };
            openLog.Click += (s, e) => OpenFolder(Log.Folder);
            var openSettings = new Button { Text = "Open settings &folder", Left = left + 136, Top = top, Width = 150, Height = 28 };
            openSettings.Click += (s, e) => OpenFolder(OutlookSettings.Folder);
            var ok = new Button { Text = "OK", Left = total - left - 186, Top = top, Width = 90, Height = 28 };
            ok.Click += (s, e) => Apply();
            var cancel = new Button { Text = "Cancel", Left = total - left - 90, Top = top, Width = 90, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { openLog, openSettings, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;

            ClientSize = new Size(total, top + 28 + 14);
        }

        private static TextBox OpeningBox(GroupBox group, string label, string value, ref int y)
        {
            group.Controls.Add(new Label { Text = label, Left = 12, Top = y + 4, Width = 120 });
            var box = new TextBox { Left = 136, Top = y, Width = group.Width - 148, Text = value ?? string.Empty };
            group.Controls.Add(box);
            y += 30;
            return box;
        }

        private void Apply()
        {
            List<string> presets = _presets.Lines.Select(l => l.Trim()).Where(l => l.Length > 0).ToList();
            List<string> bad = presets.Where(p => SendLaterTime.ParsePreset(p) == null).ToList();
            if (bad.Count > 0)
            {
                _error.Text = "Not understood: " + string.Join(", ", bad) + ". Fix or remove these lines.";
                return;
            }
            if (presets.Count == 0)
            {
                _error.Text = "Keep at least one Send later preset.";
                return;
            }

            TimeSpan? workStart = AvailabilityText.ParseTime(_workStart.Text), workEnd = AvailabilityText.ParseTime(_workEnd.Text);
            if (!workStart.HasValue || !workEnd.HasValue || workEnd.Value <= workStart.Value)
            {
                _error.Text = "Working hours: write two times such as 08:00 and 17:00, the second after the first.";
                return;
            }
            List<DayOfWeek> workDays = _dayBoxes.Where(kv => kv.Value.Checked).Select(kv => kv.Key).ToList();
            if (workDays.Count == 0)
            {
                _error.Text = "Tick at least one working day.";
                return;
            }

            _edit.AvailabilityWorkStart = workStart.Value.ToString(@"hh\:mm");
            _edit.AvailabilityWorkEnd = workEnd.Value.ToString(@"hh\:mm");
            _edit.AvailabilityWorkDays = AvailabilityText.DaysText(workDays);
            _edit.AvailabilityBuffer = (int)_buffer.Value;
            _edit.AvailabilityOpeningListEn = _openListEn.Text.Trim();
            _edit.AvailabilityOpeningListSv = _openListSv.Text.Trim();
            _edit.AvailabilityOpeningSentenceEn = _openSentEn.Text.Trim();
            _edit.AvailabilityOpeningSentenceSv = _openSentSv.Text.Trim();

            foreach (var kv in _featureBoxes)
            {
                FeatureInfo info = Feature.Find(kv.Key);
                if (info != null && info.Built) _edit.SetFeatureOn(kv.Key, kv.Value.Checked);
            }
            _edit.SendLaterPresets = presets;
            _edit.SendLaterRoomWarning = _roomWarning.Checked;
            _edit.SaveAsCurrent();
            Log.Info("Settings saved");

            DialogResult = DialogResult.OK;
        }

        private void OpenFolder(string folder)
        {
            try
            {
                Directory.CreateDirectory(folder);
                Process.Start("explorer.exe", "\"" + folder + "\"");
            }
            catch (Exception ex)
            {
                _error.Text = "Could not open " + folder + ": " + ex.Message;
            }
        }
    }
}
