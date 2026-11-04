using System;
using ProdigoOutlook.Core;
using ProdigoOutlook.Features.SendLater;
using ProdigoOutlook.OnSend;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Deliberately almost nothing happens here.
        ///
        /// Outlook times every add-in's startup and disables the ones that are slow
        /// (roughly: an average over a second), and the only thing the user sees is a
        /// yellow bar and a missing Prodigo tab. So startup only hooks two events, both
        /// cheap; settings, forms and the Outbox are all touched on first use instead.
        ///
        /// Anything thrown here makes Office hard-disable the add-in, hence the catch.
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                Application.ItemSend += Application_ItemSend;
                OutboxGuard.Start(Application);
            }
            catch (Exception ex)
            {
                Log.Error("Startup failed; Prodigo's send features are inactive this session", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Outlook no longer raises Shutdown for add-ins by default (fast shutdown),
            // so nothing important may depend on running here.
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            // The dispatcher fails open by itself; this is only the last line of defence.
            try
            {
                OnSendDispatcher.OnItemSend(item, ref cancel);
            }
            catch (Exception ex)
            {
                Log.Error("ItemSend: unexpected failure; item is sent as normal", ex);
            }
        }

        internal Ribbon RibbonInstance { get; private set; }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            RibbonInstance = new Ribbon();
            return RibbonInstance;
        }

        /// <summary>
        /// The object AutoHotkey reaches through
        /// ComObjActive("Outlook.Application").COMAddIns.Item("ProdigoOutlook").Object -
        /// the same bridge, and the same three operations, as the PowerPoint add-in.
        /// </summary>
        private AhkBridge bridge;
        protected override object RequestComAddInAutomationService()
            => bridge ?? (bridge = new AhkBridge());

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
