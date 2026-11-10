using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// What the user has chosen, saved as key = value lines in
    /// %APPDATA%\Prodigo\outlook-settings.txt.
    ///
    /// The same shape as ProdigoExcel's MapSettings, on purpose: a file rather than the
    /// registry or .NET user settings, so it can be read, edited by hand, backed up and
    /// copied to another computer, and so "send me your settings file" is a support
    /// answer that works. (The PowerPoint add-in's .NET settings reset themselves on
    /// every version bump; this file does not care about versions.) Named for Outlook
    /// because the other add-ins keep their own files in the same folder.
    ///
    /// Nothing here throws. A missing file means defaults, and so does a corrupt one.
    /// Saving goes via a .tmp file and File.Replace, so an interrupted write cannot
    /// truncate the real file.
    ///
    /// Loaded on first use, not at Outlook startup, to keep startup fast.
    ///
    /// Keys:
    ///   feature.&lt;id&gt;          = on | off      (see Features)
    ///   sendlater.preset       = tomorrow 08:00 (repeated, one per preset, in order)
    ///   sendlater.roomwarning  = on | off
    ///   availability.*         = working hours and days, buffer, opening lines, and
    ///                            the last choices in the window (see AvailabilityKeys)
    /// The first sendlater.preset line replaces the default list rather than adding to
    /// it, so a file without any means the defaults.
    /// </summary>
    internal sealed class OutlookSettings
    {
        private static OutlookSettings _current;

        public static string Folder => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Prodigo");

        public static string FilePath => Path.Combine(Folder, "outlook-settings.txt");

        public static OutlookSettings Current => _current ?? (_current = Load());

        private readonly Dictionary<string, bool> _features = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        public List<string> SendLaterPresets { get; set; } = new List<string>(Defaults.SendLaterPresets);

        public bool SendLaterRoomWarning { get; set; } = Defaults.SendLaterRoomWarning;

        // Insert availability. Times as HH:mm, days as "mon,tue,...", buffer in minutes.
        public string AvailabilityWorkStart { get; set; } = Defaults.AvailabilityWorkStart;
        public string AvailabilityWorkEnd { get; set; } = Defaults.AvailabilityWorkEnd;
        public string AvailabilityWorkDays { get; set; } = Defaults.AvailabilityWorkDays;
        public int AvailabilityBuffer { get; set; } = Defaults.AvailabilityBuffer;
        public string AvailabilityOpeningListEn { get; set; } = Defaults.AvailabilityOpeningListEn;
        public string AvailabilityOpeningListSv { get; set; } = Defaults.AvailabilityOpeningListSv;
        public string AvailabilityOpeningSentenceEn { get; set; } = Defaults.AvailabilityOpeningSentenceEn;
        public string AvailabilityOpeningSentenceSv { get; set; } = Defaults.AvailabilityOpeningSentenceSv;

        // The last choices in the Insert availability window, remembered between uses.
        public int AvailabilityLastLength { get; set; } = 60;
        public string AvailabilityLastPeriod { get; set; } = "thisweek";
        public string AvailabilityLastFormat { get; set; } = "list";
        public string AvailabilityLastLanguage { get; set; } = "sv";

        private static readonly (string Key, Func<OutlookSettings, string> Get, Action<OutlookSettings, string> Set)[] AvailabilityKeys =
        {
            ("availability.workstart", s => s.AvailabilityWorkStart, (s, v) => s.AvailabilityWorkStart = v),
            ("availability.workend", s => s.AvailabilityWorkEnd, (s, v) => s.AvailabilityWorkEnd = v),
            ("availability.workdays", s => s.AvailabilityWorkDays, (s, v) => s.AvailabilityWorkDays = v),
            ("availability.buffer", s => s.AvailabilityBuffer.ToString(), (s, v) => { int n; if (int.TryParse(v, out n) && n >= 0 && n <= 240) s.AvailabilityBuffer = n; }),
            ("availability.opening.list.en", s => s.AvailabilityOpeningListEn, (s, v) => s.AvailabilityOpeningListEn = v),
            ("availability.opening.list.sv", s => s.AvailabilityOpeningListSv, (s, v) => s.AvailabilityOpeningListSv = v),
            ("availability.opening.sentence.en", s => s.AvailabilityOpeningSentenceEn, (s, v) => s.AvailabilityOpeningSentenceEn = v),
            ("availability.opening.sentence.sv", s => s.AvailabilityOpeningSentenceSv, (s, v) => s.AvailabilityOpeningSentenceSv = v),
            ("availability.last.length", s => s.AvailabilityLastLength.ToString(), (s, v) => { int n; if (int.TryParse(v, out n) && n > 0) s.AvailabilityLastLength = n; }),
            ("availability.last.period", s => s.AvailabilityLastPeriod, (s, v) => s.AvailabilityLastPeriod = v),
            ("availability.last.format", s => s.AvailabilityLastFormat, (s, v) => s.AvailabilityLastFormat = v),
            ("availability.last.language", s => s.AvailabilityLastLanguage, (s, v) => s.AvailabilityLastLanguage = v),
        };

        /// <summary>Built features are on unless the file says otherwise.</summary>
        public bool IsFeatureOn(string id)
        {
            bool on;
            return !_features.TryGetValue(id, out on) || on;
        }

        public void SetFeatureOn(string id, bool on) => _features[id] = on;

        public OutlookSettings Copy()
        {
            var copy = (OutlookSettings)MemberwiseClone();
            copy.SendLaterPresets = new List<string>(SendLaterPresets);
            foreach (var kv in _features) copy._features[kv.Key] = kv.Value;
            return copy;
        }

        /// <summary>Makes this the settings in force and writes it to disk.</summary>
        public void SaveAsCurrent()
        {
            _current = this;
            try
            {
                Directory.CreateDirectory(Folder);
                var sb = new StringBuilder();
                sb.AppendLine("# Prodigo for Outlook settings. Edit with Outlook closed, or use Prodigo > Settings.");
                foreach (FeatureInfo f in Feature.All)
                {
                    bool on;
                    if (_features.TryGetValue(f.Id, out on))
                        sb.AppendLine("feature." + f.Id + " = " + (on ? "on" : "off"));
                }
                foreach (string preset in SendLaterPresets)
                    sb.AppendLine("sendlater.preset = " + preset);
                sb.AppendLine("sendlater.roomwarning = " + (SendLaterRoomWarning ? "on" : "off"));
                foreach (var k in AvailabilityKeys)
                    sb.AppendLine(k.Key + " = " + (k.Get(this) ?? string.Empty));

                string tmp = FilePath + ".tmp";
                File.WriteAllText(tmp, sb.ToString(), Encoding.UTF8);
                if (File.Exists(FilePath))
                    File.Replace(tmp, FilePath, null);
                else
                    File.Move(tmp, FilePath);
            }
            catch (Exception ex)
            {
                Log.Error("Could not save settings", ex);
            }
        }

        private static OutlookSettings Load()
        {
            var s = new OutlookSettings();
            try
            {
                if (!File.Exists(FilePath)) return s;

                bool presetsReplaced = false;
                foreach (string raw in File.ReadAllLines(FilePath))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith("#")) continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim().ToLowerInvariant();
                    string value = line.Substring(eq + 1).Trim();

                    if (key.StartsWith("feature."))
                    {
                        s._features[key.Substring("feature.".Length)] = IsOn(value);
                    }
                    else if (key == "sendlater.preset")
                    {
                        if (!presetsReplaced)
                        {
                            s.SendLaterPresets.Clear();
                            presetsReplaced = true;
                        }
                        if (value.Length > 0) s.SendLaterPresets.Add(value);
                    }
                    else if (key == "sendlater.roomwarning")
                    {
                        s.SendLaterRoomWarning = IsOn(value);
                    }
                    else
                    {
                        foreach (var k in AvailabilityKeys)
                            if (k.Key == key) k.Set(s, value);
                    }
                }

                if (s.SendLaterPresets.Count == 0)
                    s.SendLaterPresets = new List<string>(Defaults.SendLaterPresets);
            }
            catch (Exception ex)
            {
                Log.Error("Could not read settings; using defaults", ex);
                return new OutlookSettings();
            }
            return s;
        }

        private static bool IsOn(string value) =>
            new[] { "on", "true", "yes", "1" }.Contains(value.Trim().ToLowerInvariant());
    }
}
