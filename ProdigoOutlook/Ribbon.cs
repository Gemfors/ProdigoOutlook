using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ProdigoOutlook.Core;
using ProdigoOutlook.Features.SendLater;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook
{
    /// <summary>
    /// The Prodigo tab in every Outlook window, and one callback per button.
    ///
    /// Unlike PowerPoint, Outlook asks for its ribbon once per *kind* of window - the
    /// main window, a compose window, a read window, an appointment - and passes an ID
    /// saying which. So there is one XML file per window type in Ribbon\, each showing
    /// only what makes sense there, and GetCustomUI picks one.
    ///
    /// The ribbon XML is all-or-nothing, as CLAUDE.md in the PowerPoint project warns:
    /// one attribute Office doesn't recognise and the whole tab silently disappears.
    /// If the tab is missing after a build, turn on File > Options > Advanced >
    /// Developers > "Show add-in user interface errors" first.
    ///
    /// Every command callback is public, returns void and takes one IRibbonControl -
    /// the shape AhkBridge looks for by reflection. The [Feature] attribute says which
    /// feature a command belongs to, so the bridge lists only what is switched on.
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        internal static readonly Dictionary<string, string> ResourcesByRibbonId = new Dictionary<string, string>
        {
            { "Microsoft.Outlook.Explorer",     "ProdigoOutlook.Ribbon.Explorer.xml" },
            { "Microsoft.Outlook.Mail.Compose", "ProdigoOutlook.Ribbon.MailCompose.xml" },
            { "Microsoft.Outlook.Mail.Read",    "ProdigoOutlook.Ribbon.MailRead.xml" },
            { "Microsoft.Outlook.Appointment",  "ProdigoOutlook.Ribbon.Appointment.xml" },
        };

        // One IRibbonUI per window type. Invalidate refreshes every window of that type.
        private readonly List<Office.IRibbonUI> _ribbons = new List<Office.IRibbonUI>();

        public string GetCustomUI(string ribbonID)
        {
            string resource;
            return ResourcesByRibbonId.TryGetValue(ribbonID, out resource) ? ReadResource(resource) : null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbons.Add(ribbonUI);
        }

        /// <summary>Re-asks every visibility question, e.g. after Settings switched a feature off.</summary>
        internal void InvalidateAll()
        {
            foreach (Office.IRibbonUI ui in _ribbons)
            {
                try { ui.Invalidate(); } catch { /* window type closed */ }
            }
        }

        /// <summary>
        /// Shown when the control's feature is built and switched on. The tag holds the
        /// feature id - or several, comma-separated, for a group that should show when
        /// any of its buttons does - optionally followed by "|organizer" to hide the
        /// control on meetings you were invited to rather than organised.
        /// </summary>
        public bool GetVisible(Office.IRibbonControl control)
        {
            try
            {
                string tag = control.Tag ?? string.Empty;
                string[] parts = tag.Split('|');
                bool anyOn = parts[0].Split(',').Any(id => Feature.IsOn(id.Trim()));
                if (!anyOn) return false;

                if (parts.Length > 1 && parts[1] == "organizer")
                    return !IsAttendeeView(control);

                return true;
            }
            catch (Exception ex)
            {
                Log.Warn("GetVisible failed for " + control?.Id, ex);
                return false;
            }
        }

        private static bool IsAttendeeView(Office.IRibbonControl control)
        {
            Outlook.Inspector inspector = control.Context as Outlook.Inspector;
            if (inspector == null) return false;

            // The window's own item: not released, the inspector is still using it.
            if (inspector.CurrentItem is Outlook.AppointmentItem appt)
            {
                Outlook.OlMeetingStatus status = appt.MeetingStatus;
                return status == Outlook.OlMeetingStatus.olMeetingReceived
                    || status == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled;
            }
            return false;
        }

        // -------------------------------------------------------------------
        // Send later (phase 1)
        // -------------------------------------------------------------------

        [ForFeature(Feature.SendLater)]
        public void OnSendLater(Office.IRibbonControl control) =>
            Ol.Run(Feature.SendLater, "Send Later", () => SendLaterCommands.SendLater(control));

        [ForFeature(Feature.SendLater)]
        public void OnScheduled(Office.IRibbonControl control) =>
            Ol.Run(Feature.SendLater, "Scheduled Sends", () => ScheduledForm.Open(control));

        // -------------------------------------------------------------------
        // Prodigo
        // -------------------------------------------------------------------

        public void OnSettings(Office.IRibbonControl control) =>
            Ol.Run(null, "Settings", () => SettingsForm.Open(control));

        /// <summary>
        /// The placeholder action for buttons of features not built yet. They are hidden
        /// until their feature is built, so this is only reached if one is shown by
        /// mistake. Built features get their own callback, which is what AutoHotkey
        /// calls - hence [ForFeature(null)] keeps this one off the bridge's list.
        /// </summary>
        [ForFeature(null)]
        public void OnPlanned(Office.IRibbonControl control)
        {
            string feature = control?.Tag ?? string.Empty;
            FeatureInfo info = Feature.Find(feature.Split('|')[0].Split(',')[0]);
            Ol.Info(info == null
                ? "This command is not built yet."
                : info.Name + " is planned for phase " + info.Phase + " and not built yet.");
        }

        private static string ReadResource(string name)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (stream == null) return null;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }

    /// <summary>
    /// Which feature a ribbon command belongs to. The AutoHotkey bridge only lists and
    /// runs commands whose feature is switched on; null means "never on the bridge".
    /// A command without the attribute (Settings) is always available.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    internal sealed class ForFeatureAttribute : Attribute
    {
        public ForFeatureAttribute(string featureId) { FeatureId = featureId; }
        public string FeatureId { get; }
    }
}
