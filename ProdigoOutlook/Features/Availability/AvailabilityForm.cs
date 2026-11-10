using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ProdigoOutlook.Core;

namespace ProdigoOutlook.Features.Availability
{
    /// <summary>
    /// The Insert availability window: meeting length, period, list or sentence, and
    /// language, with a preview of exactly what will be inserted that updates as the
    /// choices change. The calendar is read once before the window opens; changing a
    /// choice only recalculates.
    /// </summary>
    internal sealed class AvailabilityForm : Form
    {
        private readonly List<TimeRange> _busy;
        private readonly AvailabilityRules _rules;
        private readonly DateTime _now;
        private readonly OutlookSettings _settings;

        private ComboBox _length, _period;
        private RadioButton _list, _sentence, _swedish, _english;
        private TextBox _preview;
        private Button _insert;

        private AvailabilityChoice _current;

        private AvailabilityForm(List<TimeRange> busy, AvailabilityRules rules, DateTime now,
            AvailabilityLanguage language, OutlookSettings settings)
        {
            _busy = busy;
            _rules = rules;
            _now = now;
            _settings = settings;
            FormStyle.ApplyDialog(this, "Insert Availability");
            Build(language);
            Recalculate();
        }

        /// <summary>Shows the window. Null means the user cancelled.</summary>
        public static AvailabilityChoice Ask(IWin32Window owner, List<TimeRange> busy, AvailabilityRules rules,
            DateTime now, AvailabilityLanguage language, OutlookSettings settings)
        {
            using (var form = new AvailabilityForm(busy, rules, now, language, settings))
                return form.ShowDialog(owner) == DialogResult.OK ? form._current : null;
        }

        private void Build(AvailabilityLanguage language)
        {
            const int left = 16, width = 420, labelWidth = 90;
            int top = 14;

            Controls.Add(new Label { Text = "Meeting &length:", Left = left, Top = top + 4, Width = labelWidth + 20 });
            _length = new ComboBox { Left = left + labelWidth + 24, Top = top, Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (int m in AvailabilityText.Lengths) _length.Items.Add(LengthLabel(m));
            int lastLength = Array.IndexOf(AvailabilityText.Lengths, _settings.AvailabilityLastLength);
            _length.SelectedIndex = lastLength >= 0 ? lastLength : Array.IndexOf(AvailabilityText.Lengths, 60);
            Controls.Add(_length);
            top += 32;

            Controls.Add(new Label { Text = "&Period:", Left = left, Top = top + 4, Width = labelWidth + 20 });
            _period = new ComboBox { Left = left + labelWidth + 24, Top = top, Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _period.Items.AddRange(new object[] { "This week", "Next week", "Next 2 weeks" });
            _period.SelectedIndex = (int)Availability.PeriodFromKey(_settings.AvailabilityLastPeriod);
            Controls.Add(_period);
            top += 34;

            var format = new GroupBox { Text = "Format", Left = left, Top = top, Width = 200, Height = 52 };
            _list = new RadioButton { Text = "L&ist", Left = 12, Top = 22, Width = 70 };
            _sentence = new RadioButton { Text = "&Sentence", Left = 90, Top = 22, Width = 100 };
            (_settings.AvailabilityLastFormat == "sentence" ? _sentence : _list).Checked = true;
            format.Controls.AddRange(new Control[] { _list, _sentence });
            Controls.Add(format);

            var lang = new GroupBox { Text = "Language", Left = left + 220, Top = top, Width = 200, Height = 52 };
            _swedish = new RadioButton { Text = "Sv&enska", Left = 12, Top = 22, Width = 80 };
            _english = new RadioButton { Text = "E&nglish", Left = 100, Top = 22, Width = 90 };
            (language == AvailabilityLanguage.English ? _english : _swedish).Checked = true;
            lang.Controls.AddRange(new Control[] { _swedish, _english });
            Controls.Add(lang);
            top += 62;

            Controls.Add(new Label { Text = "Preview:", Left = left, Top = top, Width = width });
            top += 20;
            _preview = new TextBox
            {
                Left = left, Top = top, Width = width, Height = 150,
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                BackColor = SystemColors.Window, TabStop = false,
            };
            Controls.Add(_preview);
            top += 160;

            _insert = new Button { Text = "Insert", Left = left + width - 186, Top = top, Width = 90, Height = 28 };
            _insert.Click += (s, e) => { if (_current != null) DialogResult = DialogResult.OK; };
            var cancel = new Button { Text = "Cancel", Left = left + width - 90, Top = top, Width = 90, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { _insert, cancel });
            AcceptButton = _insert;
            CancelButton = cancel;

            ClientSize = new Size(left * 2 + width, top + 28 + 14);

            EventHandler changed = (s, e) => Recalculate();
            _length.SelectedIndexChanged += changed;
            _period.SelectedIndexChanged += changed;
            foreach (RadioButton r in new[] { _list, _sentence, _swedish, _english })
                r.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) Recalculate(); };

            ActiveControl = _insert;
        }

        private static string LengthLabel(int minutes) =>
            minutes < 60 ? minutes + " min" : minutes % 60 == 0 ? minutes / 60 + " h" : (minutes / 60.0).ToString("0.#") + " h";

        private void Recalculate()
        {
            int length = AvailabilityText.Lengths[Math.Max(0, _length.SelectedIndex)];
            var period = (AvailabilityPeriod)Math.Max(0, _period.SelectedIndex);
            AvailabilityFormat format = _sentence.Checked ? AvailabilityFormat.Sentence : AvailabilityFormat.List;
            AvailabilityLanguage language = _english.Checked ? AvailabilityLanguage.English : AvailabilityLanguage.Swedish;

            DateTime first, last;
            AvailabilityText.PeriodDays(period, _now, out first, out last);
            List<TimeRange> free = AvailabilityText.FreeRanges(_busy, _rules, first, last, _now, TimeSpan.FromMinutes(length));
            List<TimeRange> picked = AvailabilityText.Pick(free);
            string opening = Opening(format, language);
            string text = AvailabilityText.Format(picked, format, language, opening);

            if (text == null)
            {
                _current = null;
                _preview.Text = "No free time of " + LengthLabel(length) + " in this period, within your working hours. " +
                                "Try a shorter meeting or a longer period.";
                _insert.Enabled = false;
                return;
            }

            _current = new AvailabilityChoice
            {
                Text = text,
                TextWithoutBullets = AvailabilityText.Format(picked, format, language, opening, string.Empty),
                HasOpening = !string.IsNullOrWhiteSpace(opening),
                RangeCount = picked.Count,
                LengthMinutes = length,
                Period = period,
                Format = format,
                Language = language,
            };
            _preview.Text = text.Replace("\n", Environment.NewLine);
            _insert.Enabled = true;
        }

        private string Opening(AvailabilityFormat format, AvailabilityLanguage language)
        {
            bool sv = language == AvailabilityLanguage.Swedish;
            return format == AvailabilityFormat.List
                ? (sv ? _settings.AvailabilityOpeningListSv : _settings.AvailabilityOpeningListEn)
                : (sv ? _settings.AvailabilityOpeningSentenceSv : _settings.AvailabilityOpeningSentenceEn);
        }
    }
}
