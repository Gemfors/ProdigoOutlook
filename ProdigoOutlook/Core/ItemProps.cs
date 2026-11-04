using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Reads and writes Prodigo's own hidden notes on Outlook items - the "stamps".
    ///
    /// Stored as named MAPI properties in a property set only Prodigo uses, through
    /// PropertyAccessor, rather than as Outlook UserProperties. The difference is
    /// invisible day to day but matters on email: an item carrying UserProperties is
    /// treated as a custom form, which is a classic cause of external recipients
    /// getting attachments as an unreadable winmail.dat. Named properties have no such
    /// side effect, never appear in Outlook's field chooser, and cannot collide with
    /// anything else because the property set GUID below is Prodigo's alone.
    ///
    /// Values are always strings (times via TimeHelpers.ToStored, in UTC). That avoids
    /// PropertyAccessor's own date conversion rules, and keeps a value readable when
    /// someone inspects an item with a MAPI tool.
    ///
    /// Reading a property that is not there makes Outlook throw, so Get returns null
    /// for "not set" and for any other failure alike. Nothing here throws.
    /// </summary>
    internal static class ItemProps
    {
        /// <summary>Prodigo's property set. Never change it: existing stamps would become invisible.</summary>
        private const string PropertySet = "{AE5A16EB-D428-4C1F-A3F7-ED827506DCD7}";

        private const string NamedPrefix = "http://schemas.microsoft.com/mapi/string/" + PropertySet + "/";

        /// <summary>PR_DEFERRED_SEND_TIME: when the store submits a deferred message. UTC.</summary>
        public const string DeferredSendTime = "http://schemas.microsoft.com/mapi/proptag/0x3FEF0040";

        /// <summary>PR_DEFERRED_DELIVERY_TIME: what the object model's DeferredDeliveryTime maps to. UTC.</summary>
        public const string DeferredDeliveryTime = "http://schemas.microsoft.com/mapi/proptag/0x000F0040";

        /// <summary>PR_MESSAGE_FLAGS. Bit 0x4 (MSGFLAG_SUBMIT) means "queued to be sent".</summary>
        public const string MessageFlags = "http://schemas.microsoft.com/mapi/proptag/0x0E070003";

        public const int MsgFlagSubmit = 0x4;

        /// <summary>PR_DISPLAY_TYPE_EX on an address entry; low byte 7 (DT_ROOM) is a meeting room.</summary>
        public const string DisplayTypeEx = "http://schemas.microsoft.com/mapi/proptag/0x39050003";

        /// <summary>
        /// The full schema name of one of Prodigo's notes, typed as a string, for reading
        /// it from a folder's table of contents (Folder.GetTable) or searching by it.
        /// The "/0x0000001F" (PT_UNICODE) suffix is required there: as a table column
        /// the untyped name is refused outright, and in a search filter it is accepted
        /// but silently matches nothing. PropertyAccessor (Get/Set below) does not need
        /// it - it takes the type from the value.
        /// </summary>
        public static string TableSchema(string name) => NamedPrefix + name + "/0x0000001F";

        public static string Get(object item, string name)
        {
            Outlook.PropertyAccessor pa = Accessor(item);
            if (pa == null) return null;
            try
            {
                object value = pa.GetProperty(NamedPrefix + name);
                return value as string;
            }
            catch
            {
                return null;
            }
            finally
            {
                ComScope.Release(pa);
            }
        }

        public static bool Set(object item, string name, string value)
        {
            Outlook.PropertyAccessor pa = Accessor(item);
            if (pa == null) return false;
            try
            {
                if (value == null)
                {
                    try { pa.DeleteProperty(NamedPrefix + name); } catch { /* was not there */ }
                }
                else
                {
                    pa.SetProperty(NamedPrefix + name, value);
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Warn("Could not write property " + name, ex);
                return false;
            }
            finally
            {
                ComScope.Release(pa);
            }
        }

        public static DateTime? GetTime(object item, string name) => TimeHelpers.FromStored(Get(item, name));

        public static bool SetTime(object item, string name, DateTime? local) =>
            Set(item, name, local.HasValue ? TimeHelpers.ToStored(local.Value) : null);

        /// <summary>A raw MAPI property by schema name, or null.</summary>
        public static object GetRaw(object item, string schemaName)
        {
            Outlook.PropertyAccessor pa = Accessor(item);
            if (pa == null) return null;
            try
            {
                return pa.GetProperty(schemaName);
            }
            catch
            {
                return null;
            }
            finally
            {
                ComScope.Release(pa);
            }
        }

        /// <summary>
        /// True when the item is queued to be sent. An item in the Outbox without this
        /// flag will sit there forever - which is what opening a scheduled item does.
        /// </summary>
        public static bool IsSubmitted(object item)
        {
            object flags = GetRaw(item, MessageFlags);
            return flags is int && (((int)flags) & MsgFlagSubmit) != 0;
        }

        /// <summary>
        /// Sets the deferral directly in MAPI. The fallback for when the object model's
        /// DeferredDeliveryTime does not stick. Both properties are written, since Outlook
        /// may recompute one from the other on submit. PropertyAccessor wants UTC.
        /// </summary>
        public static bool SetDeferredSendTimeRaw(object item, DateTime local)
        {
            Outlook.PropertyAccessor pa = Accessor(item);
            if (pa == null) return false;
            try
            {
                DateTime utc = local.ToUniversalTime();
                pa.SetProperty(DeferredDeliveryTime, utc);
                pa.SetProperty(DeferredSendTime, utc);
                return true;
            }
            catch (Exception ex)
            {
                Log.Warn("Could not set PR_DEFERRED_SEND_TIME", ex);
                return false;
            }
            finally
            {
                ComScope.Release(pa);
            }
        }

        private static Outlook.PropertyAccessor Accessor(object item)
        {
            try
            {
                if (item is Outlook.MailItem mail) return mail.PropertyAccessor;
                if (item is Outlook.MeetingItem meeting) return meeting.PropertyAccessor;
                if (item is Outlook.AppointmentItem appt) return appt.PropertyAccessor;
                if (item is Outlook.AddressEntry entry) return entry.PropertyAccessor;
            }
            catch
            {
                // Item already gone, or not an item that supports it.
            }
            return null;
        }
    }
}
