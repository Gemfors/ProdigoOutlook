using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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

            _error = new Label { Left = left, Top = top, Width = width, Height = 34, ForeColor = Color.Firebrick };
            Controls.Add(_error);
            top += 38;

            var openLog = new Button { Text = "Open &log folder", Left = left, Top = top, Width = 130, Height = 28 };
            openLog.Click += (s, e) => OpenFolder(Log.Folder);
            var openSettings = new Button { Text = "Open settings &folder", Left = left + 136, Top = top, Width = 150, Height = 28 };
            openSettings.Click += (s, e) => OpenFolder(OutlookSettings.Folder);
            var ok = new Button { Text = "OK", Left = left + width - 186, Top = top, Width = 90, Height = 28 };
            ok.Click += (s, e) => Apply();
            var cancel = new Button { Text = "Cancel", Left = left + width - 90, Top = top, Width = 90, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { openLog, openSettings, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;

            ClientSize = new Size(left * 2 + width, top + 28 + 14);
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
