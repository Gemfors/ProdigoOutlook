using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using ProdigoOutlook.Core;
using Office = Microsoft.Office.Core;

namespace ProdigoOutlook
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAhkBridge
    {
        string Ping();
        string Invoke(string command);
        string ListCommands();
    }

    /// <summary>
    /// The AutoHotkey bridge - the same three operations as the PowerPoint add-in's, so
    /// all three Prodigo add-ins look alike from a script:
    /// <code>
    /// svc := ComObjActive("Outlook.Application").COMAddIns.Item("ProdigoOutlook").Object
    /// svc.Ping()                 ; version and window, or an error
    /// svc.ListCommands()         ; newline-separated command names
    /// svc.Invoke("Send Later")   ; "OK", or "ERROR: ..."
    /// </code>
    /// This is the reliable way to put a command on a hotkey. Keytips (Alt, Z, P, S, L)
    /// work too, but keystroke sequences are timing-sensitive - the AHK notes record a
    /// sequence losing its Alt and typing letters into a cell.
    ///
    /// A command runs exactly as if its button had been clicked, in the Outlook window
    /// that is in front: Send Later in a compose window schedules that email. "OK"
    /// means dispatched; a command that opens a form blocks the call until the form is
    /// closed.
    ///
    /// Commands are found by reflection over Ribbon (public, void, one IRibbonControl),
    /// listed under their screentips, and filtered by their [Feature] - so a new button
    /// appears here with no change to this file, and a switched-off feature disappears.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AhkBridge : StandardOleMarshalObject, IAhkBridge
    {
        private static readonly string[] BuiltInCommands = { "Ping", "ListCommands" };

        // Fixed for the life of the add-in: the XML is compiled in and so are the
        // callbacks. COM calls arrive on Outlook's thread (StandardOleMarshalObject
        // marshals them there), so no locking is needed.
        private static Dictionary<string, string> _screentipsByMethod;

        public string Ping()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                object window = app.ActiveWindow();
                string kind = window is Microsoft.Office.Interop.Outlook.Inspector ? "item window"
                            : window is Microsoft.Office.Interop.Outlook.Explorer ? "main window"
                            : "no window";
                return "ProdigoOutlook bridge OK - Outlook " + app.Version + ", in front: " + kind;
            }
            catch (Exception ex)
            {
                return "Bridge reached, but the object model threw: " + ex.Message;
            }
        }

        public string Invoke(string command)
        {
            try
            {
                command = command == null ? string.Empty : command.Trim();
                if (command.Length == 0)
                    return "ERROR: no command given. Try \"Ping\" or \"ListCommands\".";

                string key = Normalise(command);
                if (key == "ping") return Ping();
                if (key == "listcommands") return ListCommands();

                var ribbon = Globals.ThisAddIn.RibbonInstance;
                if (ribbon == null)
                    return "ERROR: ribbon not created yet";

                MethodInfo callback = Available().FirstOrDefault(m =>
                    Normalise(m.Name) == key || Normalise(FriendlyName(m)) == key);
                if (callback == null)
                    return "ERROR: unknown or switched-off command \"" + command +
                           "\". Run ListCommands to see the names.";

                callback.Invoke(ribbon, new object[] { null });
                return "OK";
            }
            catch (TargetInvocationException ex)
            {
                return "ERROR: " + (ex.InnerException ?? ex).Message;
            }
            catch (Exception ex)
            {
                return "ERROR: " + ex.Message;
            }
        }

        public string ListCommands()
        {
            try
            {
                var names = Available()
                    .Select(FriendlyName)
                    .Concat(BuiltInCommands)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase);
                return string.Join(Environment.NewLine, names);
            }
            catch (Exception ex)
            {
                return "ERROR: " + ex.Message;
            }
        }

        /// <summary>Ribbon commands whose feature is on right now, ordered by name.</summary>
        private static List<MethodInfo> Available()
        {
            return typeof(Ribbon)
                .GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.ReturnType == typeof(void))
                .Where(m =>
                {
                    var p = m.GetParameters();
                    return p.Length == 1 && p[0].ParameterType == typeof(Office.IRibbonControl);
                })
                .Where(m =>
                {
                    var feature = m.GetCustomAttribute<ForFeatureAttribute>();
                    if (feature == null) return true;
                    return feature.FeatureId != null && Feature.IsOn(feature.FeatureId);
                })
                .OrderBy(m => m.Name, StringComparer.Ordinal)
                .ToList();
        }

        private static string FriendlyName(MethodInfo callback)
        {
            string screentip;
            return ScreentipsByMethod().TryGetValue(callback.Name, out screentip) && screentip.Length > 0
                ? screentip
                : callback.Name;
        }

        /// <summary>Callback name to screentip, read from all four embedded ribbon files.</summary>
        private static Dictionary<string, string> ScreentipsByMethod()
        {
            if (_screentipsByMethod != null) return _screentipsByMethod;

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string resource in Ribbon.ResourcesByRibbonId.Values)
            {
                try
                {
                    using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
                    {
                        if (stream == null) continue;
                        foreach (XElement element in XDocument.Load(stream).Descendants())
                        {
                            XAttribute action = element.Attribute("onAction");
                            XAttribute screentip = element.Attribute("screentip");
                            if (action == null || screentip == null) continue;
                            if (action.Value.Length == 0 || screentip.Value.Length == 0) continue;
                            if (!map.ContainsKey(action.Value))
                                map[action.Value] = screentip.Value;
                        }
                    }
                }
                catch
                {
                    // A malformed file only costs the friendly names; commands still run.
                }
            }

            _screentipsByMethod = map;
            return _screentipsByMethod;
        }

        /// <summary>Letters and digits only, lower-cased: "Send Later" = "sendlater" = "OnSendLater" minus "on".</summary>
        private static string Normalise(string name)
        {
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
                if (char.IsLetterOrDigit(c)) sb.Append(char.ToLowerInvariant(c));
            string s = sb.ToString();
            return s.StartsWith("on") && s.Length > 2 ? s.Substring(2) : s;
        }
    }
}
