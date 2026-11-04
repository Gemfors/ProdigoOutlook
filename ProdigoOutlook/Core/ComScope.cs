using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Releases the Outlook objects a piece of code fetched, when it is done with them.
    ///
    /// This matters far more in Outlook than in PowerPoint or Excel. Every Outlook
    /// object .NET touches is held open until the garbage collector gets round to it,
    /// and an item that is still held open causes the familiar trouble: "the item has
    /// been changed" when you save it, an inspector that will not really close, and in
    /// online mode Exchange refusing to open more than about 255 items at once.
    ///
    /// Usage:
    /// <code>
    /// using (var com = new ComScope())
    /// {
    ///     var folder = com.Track(ns.GetDefaultFolder(...));
    ///     var items  = com.Track(folder.Items);
    /// }
    /// </code>
    ///
    /// Two rules, both learned the hard way by everyone who writes Outlook add-ins:
    /// - Only track what you fetched yourself. Never track the item Outlook passes into
    ///   an event such as ItemSend, nor Application: those are shared with Outlook and
    ///   releasing them breaks the next piece of code that uses the same wrapper.
    /// - Avoid chains like a.B.C inside loops: each dot creates an object nobody can
    ///   release. Fetch each step into a tracked variable instead.
    /// </summary>
    internal sealed class ComScope : IDisposable
    {
        private readonly List<object> _objects = new List<object>();

        public T Track<T>(T obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
                _objects.Add(obj);
            return obj;
        }

        public void Dispose()
        {
            for (int i = _objects.Count - 1; i >= 0; i--)
                Release(_objects[i]);
            _objects.Clear();
        }

        public static void Release(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // Already released, or Outlook is shutting down.
            }
        }
    }
}
