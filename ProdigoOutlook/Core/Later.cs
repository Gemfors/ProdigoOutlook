using System;
using System.Windows.Forms;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Runs a piece of work a little later, on Outlook's own thread.
    ///
    /// Several things cannot safely be done inside an Outlook event: saving the
    /// appointment while Outlook is in the middle of sending its invite, or touching an
    /// item while its window is closing. Outlook is still working on the item at that
    /// moment and will overwrite or reject the change. Doing it a second or two later,
    /// once Outlook has finished, avoids the race.
    ///
    /// A WinForms timer rather than a thread timer, deliberately: Outlook's objects may
    /// only be used from the thread Outlook runs the add-in on, and a WinForms timer
    /// ticks on the thread that created it. So <see cref="Run"/> must itself be called
    /// from Outlook's thread, which every event handler and ribbon callback is.
    /// </summary>
    internal static class Later
    {
        public static void Run(TimeSpan delay, string what, Action action)
        {
            var timer = new Timer { Interval = Math.Max(1, (int)delay.TotalMilliseconds) };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    Log.Error("Deferred task failed: " + what, ex);
                }
            };
            timer.Start();
        }
    }
}
