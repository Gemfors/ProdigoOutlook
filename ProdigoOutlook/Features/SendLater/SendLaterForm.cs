using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ProdigoOutlook.Core;

namespace ProdigoOutlook.Features.SendLater
{
    /// <summary>
    /// The Send later dialog, shared by emails and meetings, built in code like the
    /// other Prodigo dialogs.
    ///
    /// Designed to be quicker than Outlook's own Delay Delivery: a preset is one click
    /// (or one key, 1 to 9) and sends straight away. The date and time pickers below
    /// are for anything else. When the item is already scheduled the form says when,
    /// and offers Send now.
    ///
    /// Use <see cref="Ask"/>; it returns null when the user cancels.
    /// </summary>
    internal sealed class SendLaterForm : Form
    {
        private readonly List<Button> _presetButtons = new List<Button>();
        private readonly Dictionary<Button, DateTime> _presetTimes = new Dictionary<Button, DateTime>();
        private DateTimePicker _date, _time;
        private Label _error;

        private SendLaterChoice _choice;

        private SendLaterForm(string subject, DateTime? currentlyScheduled, IList<string> rooms, bool isMeeting)
        {
            FormStyle.ApplyDialog(this, isMeeting ? "Send Invite Later" : "Send Later");
            Build(subject, currentlyScheduled, rooms, isMeeting);
        }

        /// <summary>Shows the form. Null means the user cancelled.</summary>
        public static SendLaterChoice Ask(IWin32Window owner, string subject, DateTime? currentlyScheduled,
            IList<string> rooms, bool isMeeting)
        {
            using (var form = new SendLaterForm(subject, currentlyScheduled, rooms, isMeeting))
            {
                return form.ShowDialog(owner) == DialogResult.OK ? form._choice : null;
            }
        }

        private void Build(string subject, DateTime? current, IList<string> rooms, bool isMeeting)
        {
            const int left = 16, width = 380;
            int top = 14;
            DateTime now = DateTime.Now;

            var subjectLabel = new Label
            {
                Text = string.IsNullOrWhiteSpace(subject) ? "(no subject)" : subject,
                Left = left, Top = top, Width = width, Height = 20, AutoEllipsis = true,
                Font = new Font(Font, FontStyle.Bold),
            };
            Controls.Add(subjectLabel);
            top += 24;

            if (current.HasValue)
            {
                Controls.Add(new Label
                {
                    Text = "Currently scheduled: " + TimeHelpers.Friendly(current.Value) + ". Pick a new time, or send now.",
                    Left = left, Top = top, Width = width, Height = 34,
                });
                top += 38;
            }

            Controls.Add(new Label { Text = isMeeting ? "Deliver the invite:" : "Deliver:", Left = left, Top = top, Width = width, Height = 18 });
            top += 22;

            // Presets: one button each, with the actual date beside the name.
            int number = 1;
            foreach (string spec in OutlookSettings.Current.SendLaterPresets)
            {
                SendLaterPreset preset = SendLaterTime.ParsePreset(spec);
                if (preset == null) continue;

                DateTime? when = preset.Resolve(now);
                string key = number <= 9 ? "&" + number + "  " : "    ";
                var button = new Button
                {
                    Text = key + preset.Label + (when.HasValue ? "   (" + TimeHelpers.FriendlyDate(when.Value) + ")" : "   (passed)"),
                    Left = left, Top = top, Width = width, Height = 30,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Enabled = when.HasValue,
                    UseMnemonic = true,
                };
                if (when.HasValue)
                {
                    _presetTimes[button] = when.Value;
                    button.Click += (s, e) => Choose(_presetTimes[(Button)s]);
                }
                _presetButtons.Add(button);
                Controls.Add(button);
                top += 34;
                number++;
            }

            top += 8;
            Controls.Add(new Label { Text = "Other time:", Left = left, Top = top + 4, Width = 80 });

            // Never suggest a time in the past: a date before MinDate makes the picker throw.
            DateTime suggestion = TimeHelpers.TruncateToMinute(
                current.HasValue && current.Value > now.AddMinutes(1) ? current.Value : now.AddHours(1));
            _date = new DateTimePicker
            {
                Left = left + 84, Top = top, Width = 150,
                Format = DateTimePickerFormat.Custom, CustomFormat = "ddd d MMM yyyy",
                MinDate = now.Date, Value = suggestion.Date,
            };
            _time = new DateTimePicker
            {
                Left = left + 240, Top = top, Width = 70,
                Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true,
                Value = suggestion,
            };
            var sendAt = new Button { Text = "&Send", Left = left + 316, Top = top - 1, Width = 64, Height = 27 };
            sendAt.Click += (s, e) => Choose(_date.Value.Date + _time.Value.TimeOfDay);
            Controls.AddRange(new Control[] { _date, _time, sendAt });
            top += 32;

            _error = new Label { Left = left, Top = top, Width = width, Height = 18, ForeColor = Color.Firebrick };
            Controls.Add(_error);
            top += 22;

            if (rooms != null && rooms.Count > 0)
            {
                var warning = new Label
                {
                    Text = "Note: " + string.Join(", ", rooms) + " is a meeting room. It is not reserved until the " +
                           "invite is delivered, so someone else could book it before then.",
                    Left = left, Top = top, Width = width, Height = 48,
                    ForeColor = Color.DarkGoldenrod,
                };
                Controls.Add(warning);
                top += 52;
            }

            var cancel = new Button { Text = "Cancel", Width = 90, Height = 28, Top = top, Left = left + width - 90 };
            cancel.Click += (s, e) => Close();
            Controls.Add(cancel);
            CancelButton = cancel;

            if (current.HasValue)
            {
                var sendNow = new Button { Text = "Send &now", Width = 100, Height = 28, Top = top, Left = left + width - 196 };
                sendNow.Click += (s, e) =>
                {
                    _choice = SendLaterChoice.Now();
                    DialogResult = DialogResult.OK;
                };
                Controls.Add(sendNow);
            }

            ClientSize = new Size(left * 2 + width, top + 28 + 14);

            Button first = _presetButtons.FirstOrDefault(b => b.Enabled);
            if (first != null) ActiveControl = first;
        }

        private void Choose(DateTime when)
        {
            when = TimeHelpers.TruncateToMinute(when);
            string problem = SendLaterTime.Validate(when, DateTime.Now);
            if (problem != null)
            {
                _error.Text = problem;
                return;
            }

            _choice = SendLaterChoice.At(when);
            DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// Digits 1-9 pick a preset without Alt, except while typing in a picker, where a
        /// digit is part of the date or time.
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (!(ActiveControl is DateTimePicker))
            {
                int digit = -1;
                if (keyData >= Keys.D1 && keyData <= Keys.D9) digit = keyData - Keys.D1;
                else if (keyData >= Keys.NumPad1 && keyData <= Keys.NumPad9) digit = keyData - Keys.NumPad1;

                if (digit >= 0 && digit < _presetButtons.Count)
                {
                    if (_presetButtons[digit].Enabled) _presetButtons[digit].PerformClick();
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }

    /// <summary>What the user chose in the Send later form.</summary>
    internal sealed class SendLaterChoice
    {
        private SendLaterChoice() { }

        /// <summary>Send immediately, dropping any scheduled time.</summary>
        public bool SendNow { get; private set; }

        /// <summary>The delivery time, when not SendNow.</summary>
        public DateTime When { get; private set; }

        public static SendLaterChoice Now() => new SendLaterChoice { SendNow = true };

        public static SendLaterChoice At(DateTime when) => new SendLaterChoice { When = when };
    }
}
