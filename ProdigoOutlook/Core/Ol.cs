using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Shared plumbing for the ribbon commands, the Outlook counterpart of
    /// ProdigoExcel's Xl: access to Outlook, a safe way to run a command, and working
    /// out which item and window a command is about.
    /// </summary>
    internal static class Ol
    {
        public static Outlook.Application App => Globals.ThisAddIn.Application;

        /// <summary>
        /// Runs a ribbon command, turning any exception into a message box rather than
        /// letting it escape into Outlook (which discards it, so the button would just
        /// look broken). A switched-off feature says so instead of running.
        /// </summary>
        public static void Run(string featureId, string commandName, Action action)
        {
            try
            {
                if (featureId != null && !Feature.IsOn(featureId))
                {
                    Info(commandName + " is switched off in Prodigo > Settings.", commandName);
                    return;
                }
                action();
            }
            catch (Exception ex)
            {
                Log.Error("Command failed: " + commandName, ex);
                MessageBox.Show(
                    OwnerWindow(),
                    "The command could not be completed.\n\n" + ex.Message +
                    "\n\nIf this keeps happening, the log has the details (Prodigo > Settings > Open log folder).",
                    "Prodigo - " + commandName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        public static void Info(string message, string title = null) =>
            MessageBox.Show(OwnerWindow(), message, title == null ? "Prodigo" : "Prodigo - " + title,
                MessageBoxButtons.OK, MessageBoxIcon.Information);

        /// <summary>
        /// The window a command is about. From the ribbon that is the window whose
        /// ribbon was clicked; from the AutoHotkey bridge there is no control, so it is
        /// whichever Outlook window is in front.
        /// </summary>
        public static object ContextWindow(Office.IRibbonControl control)
        {
            try
            {
                if (control != null && control.Context != null)
                    return control.Context;
                return App.ActiveWindow();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// The item a command is about: the open item in an inspector, or the first
        /// selected item in the main window. Null when there is none.
        /// </summary>
        public static object ContextItem(Office.IRibbonControl control)
        {
            object window = ContextWindow(control);
            try
            {
                if (window is Outlook.Inspector inspector)
                    return inspector.CurrentItem;

                if (window is Outlook.Explorer explorer)
                {
                    Outlook.Selection selection = explorer.Selection;
                    try
                    {
                        return selection.Count > 0 ? selection[1] : null;
                    }
                    finally
                    {
                        ComScope.Release(selection);
                    }
                }
            }
            catch
            {
                // Some views (e.g. Outlook Today) have no selection to speak of.
            }
            return null;
        }

        public static Outlook.Inspector ContextInspector(Office.IRibbonControl control) =>
            ContextWindow(control) as Outlook.Inspector;

        /// <summary>
        /// Owner for a dialog: the Outlook window in front. Without an owner a WinForms
        /// dialog from an add-in can open behind Outlook, where it is modal, invisible,
        /// and looks exactly like a hang (ProdigoExcel learned this one).
        /// </summary>
        public static IWin32Window OwnerWindow(object window = null)
        {
            try
            {
                object w = window ?? App.ActiveWindow();
                if (w is IOleWindow ole)
                {
                    IntPtr hwnd;
                    ole.GetWindow(out hwnd);
                    if (hwnd != IntPtr.Zero)
                        return new WindowHandle(hwnd);
                }
            }
            catch
            {
                // Fall back to no owner.
            }
            return null;
        }

        /// <summary>The default Outbox. The caller releases it.</summary>
        public static Outlook.MAPIFolder Outbox()
        {
            Outlook.NameSpace ns = App.Session;
            try
            {
                return ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
            }
            finally
            {
                ComScope.Release(ns);
            }
        }

        /// <summary>Fetches an item by EntryID, or null if it has gone.</summary>
        public static object ItemById(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return null;
            Outlook.NameSpace ns = App.Session;
            try
            {
                return ns.GetItemFromID(entryId);
            }
            catch
            {
                return null;
            }
            finally
            {
                ComScope.Release(ns);
            }
        }

        /// <summary>
        /// One default folder (Outbox, Sent Items, ...) from every mailbox in the profile,
        /// tracked in the caller's scope. Stores without that folder (PST archives,
        /// public folders) are skipped.
        /// </summary>
        public static List<Outlook.MAPIFolder> AllFolders(Outlook.OlDefaultFolders which, ComScope com)
        {
            var folders = new List<Outlook.MAPIFolder>();
            Outlook.NameSpace ns = com.Track(App.Session);
            Outlook.Stores stores = com.Track(ns.Stores);
            int n = stores.Count;
            for (int i = 1; i <= n; i++)
            {
                try
                {
                    Outlook.Store store = com.Track(stores[i]);
                    folders.Add(com.Track(store.GetDefaultFolder(which)));
                }
                catch
                {
                    // This store has no such folder.
                }
            }
            return folders;
        }

        /// <summary>True when the item sits in the Outbox of the mailbox it belongs to.</summary>
        public static bool IsInOutbox(object item)
        {
            using (var com = new ComScope())
            {
                try
                {
                    Outlook.MAPIFolder parent = null;
                    if (item is Outlook.MailItem mail) parent = com.Track(mail.Parent as Outlook.MAPIFolder);
                    else if (item is Outlook.MeetingItem meeting) parent = com.Track(meeting.Parent as Outlook.MAPIFolder);
                    if (parent == null) return false;

                    Outlook.Store store = com.Track(parent.Store);
                    Outlook.MAPIFolder outbox = com.Track(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox));
                    return parent.EntryID == outbox.EntryID;
                }
                catch
                {
                    return false;
                }
            }
        }

        [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            void GetWindow(out IntPtr phwnd);
            void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        }

        private sealed class WindowHandle : IWin32Window
        {
            public WindowHandle(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }
    }
}
