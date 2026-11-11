using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using ProdigoOutlook.Core;

namespace ProdigoOutlook.Features.PreSendChecks
{
    /// <summary>
    /// The pre-send rules in %APPDATA%\Prodigo\outlook-rules.json, edited through
    /// Prodigo > Settings. JSON rather than the key = value settings file, as the brief
    /// asked, so a future add-in for the new Outlook could read the same rules.
    ///
    /// Loaded on first use. A missing file means the agreed defaults; so does a file
    /// that cannot be read, and anything missing from it keeps its default. Nothing
    /// here throws - the send path depends on it.
    ///
    /// {
    ///   "checks": { "attachment": true, "replyAll": true, "mixedRecipients": true, "dateMismatch": true },
    ///   "internalDomains": [ "hiq.se", ... ],
    ///   "replyAllLimit": 10,
    ///   "attachmentWords": [ "bifoga*", "bilaga", ... ]
    /// }
    /// </summary>
    internal static class PreSendRulesFile
    {
        private static PreSendRules _current;

        public static string FilePath => Path.Combine(OutlookSettings.Folder, "outlook-rules.json");

        public static PreSendRules Current => _current ?? (_current = Load());

        public static PreSendRules Load()
        {
            var rules = new PreSendRules();
            try
            {
                if (!File.Exists(FilePath)) return rules;
                var root = new JavaScriptSerializer().DeserializeObject(File.ReadAllText(FilePath)) as Dictionary<string, object>;
                if (root == null) return rules;

                if (root.TryGetValue("checks", out object c) && c is Dictionary<string, object> checks)
                {
                    rules.CheckAttachment = Bool(checks, "attachment", rules.CheckAttachment);
                    rules.CheckReplyAll = Bool(checks, "replyAll", rules.CheckReplyAll);
                    rules.CheckMixed = Bool(checks, "mixedRecipients", rules.CheckMixed);
                    rules.CheckDates = Bool(checks, "dateMismatch", rules.CheckDates);
                }
                List<string> domains = Strings(root, "internalDomains");
                if (domains != null) rules.InternalDomains = domains;
                if (root.TryGetValue("replyAllLimit", out object limit) && limit is int n && n > 0) rules.ReplyAllLimit = n;
                List<string> words = Strings(root, "attachmentWords");
                if (words != null) rules.AttachmentWords = words;
            }
            catch (Exception ex)
            {
                Log.Error("Could not read " + FilePath + "; using the default pre-send rules", ex);
                return new PreSendRules();
            }
            return rules;
        }

        /// <summary>Makes these the rules in force and writes them to disk.</summary>
        public static void Save(PreSendRules rules)
        {
            _current = rules;
            try
            {
                var sb = new StringBuilder();
                sb.Append("{\r\n");
                sb.Append("  \"checks\": {\r\n");
                sb.Append("    \"attachment\": ").Append(Js(rules.CheckAttachment)).Append(",\r\n");
                sb.Append("    \"replyAll\": ").Append(Js(rules.CheckReplyAll)).Append(",\r\n");
                sb.Append("    \"mixedRecipients\": ").Append(Js(rules.CheckMixed)).Append(",\r\n");
                sb.Append("    \"dateMismatch\": ").Append(Js(rules.CheckDates)).Append("\r\n");
                sb.Append("  },\r\n");
                sb.Append("  \"internalDomains\": ").Append(Js(rules.InternalDomains)).Append(",\r\n");
                sb.Append("  \"replyAllLimit\": ").Append(rules.ReplyAllLimit).Append(",\r\n");
                sb.Append("  \"attachmentWords\": ").Append(Js(rules.AttachmentWords)).Append("\r\n");
                sb.Append("}\r\n");

                Directory.CreateDirectory(OutlookSettings.Folder);
                string tmp = FilePath + ".tmp";
                File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(false));
                if (File.Exists(FilePath)) File.Replace(tmp, FilePath, null);
                else File.Move(tmp, FilePath);
            }
            catch (Exception ex)
            {
                Log.Error("Could not save the pre-send rules", ex);
            }
        }

        private static bool Bool(Dictionary<string, object> d, string key, bool fallback) =>
            d.TryGetValue(key, out object v) && v is bool b ? b : fallback;

        private static List<string> Strings(Dictionary<string, object> d, string key)
        {
            if (!d.TryGetValue(key, out object v) || !(v is IEnumerable list) || v is string) return null;
            return list.OfType<string>().Select(s => s.Trim()).Where(s => s.Length > 0).ToList();
        }

        private static string Js(bool b) => b ? "true" : "false";

        private static string Js(IEnumerable<string> list) =>
            "[ " + string.Join(", ", list.Select(s => new JavaScriptSerializer().Serialize(s))) + " ]";
    }
}
