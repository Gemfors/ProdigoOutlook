using System;
using System.IO;
using System.Text;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// A plain daily text log in %LOCALAPPDATA%\Prodigo\Logs\outlook-yyyy-MM-dd.log.
    ///
    /// The PowerPoint and Excel add-ins have no log because every error there ends in
    /// a message box. Outlook is different: code that runs while an item is being sent
    /// must fail silently and let the item go, so the log is the only place such an
    /// error can be seen at all. "Send me today's log" is the support answer.
    ///
    /// Nothing here throws. A log that cannot be written must never be why an email
    /// is not sent. Files older than <see cref="KeepDays"/> are removed the first time
    /// something is logged in a session.
    /// </summary>
    internal static class Log
    {
        private const int KeepDays = 14;
        private static readonly object Gate = new object();
        private static bool _purged;

        public static string Folder => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Prodigo\Logs");

        public static string TodayFile => Path.Combine(Folder, "outlook-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");

        public static void Info(string message) => Write("INFO ", message, null);

        public static void Warn(string message, Exception ex = null) => Write("WARN ", message, ex);

        public static void Error(string message, Exception ex = null) => Write("ERROR", message, ex);

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                var sb = new StringBuilder();
                sb.Append(DateTime.Now.ToString("HH:mm:ss.fff")).Append(' ').Append(level).Append(' ').Append(message);
                if (ex != null)
                    sb.AppendLine().Append("      ").Append(ex.GetType().Name).Append(": ").Append(ex.Message)
                      .AppendLine().Append(ex.StackTrace);
                sb.AppendLine();

                lock (Gate)
                {
                    Directory.CreateDirectory(Folder);
                    File.AppendAllText(TodayFile, sb.ToString(), Encoding.UTF8);
                    if (!_purged)
                    {
                        _purged = true;
                        Purge();
                    }
                }
            }
            catch
            {
                // Disk full, folder locked, whatever: logging is best effort.
            }
        }

        private static void Purge()
        {
            try
            {
                DateTime cutoff = DateTime.Now.Date.AddDays(-KeepDays);
                foreach (string file in Directory.GetFiles(Folder, "outlook-*.log"))
                {
                    if (File.GetLastWriteTime(file) < cutoff)
                        File.Delete(file);
                }
            }
            catch
            {
                // Old logs simply stay.
            }
        }
    }
}
