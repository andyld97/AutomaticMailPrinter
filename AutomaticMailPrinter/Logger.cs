using System;
using System.Globalization;

namespace AutomaticMailPrinter
{
    public class Logger
    {
        public static readonly string LOG_PATH = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "AutomaticMailprinterLog.log");
        private readonly static LogType CurrentLogLevel = LogType.Info | LogType.Warning | LogType.Error;

        public static void LogDebug(string message = "", Exception e = null)
        {
            Log(message, e, LogType.Debug, false);
        }

        public static void LogInfo(string message = "", Exception e = null, bool sendWebHook = false)
        {
            // Info can be sended to webhook optinally, but Warnings and Errors are all sended automatically!
            Log(message, e, LogType.Info, sendWebHook);
        }

        public static void LogWarning(string message = "", Exception e = null)
        {
            Log(message, e, LogType.Warning, true);
        }

        public static void LogError(string message = "", Exception e = null)
        {
            Log(message, e, LogType.Error, true);
        }

        private static void Log(string message, Exception e, LogType type, bool sendWebHook)
        {
            // Check if log level matches with current log level - otherwise ignroe
            if ((CurrentLogLevel & type) != type)
                return;

            // Prevent logging empty messages
            if (string.IsNullOrEmpty(message) && e == null)
                return;

            // Generate log message
            var now = DateTime.Now;

            string logContent = string.Empty;
            if (!string.IsNullOrEmpty(message))
                logContent = message;

            if (e != null)
            {
                if (string.IsNullOrEmpty(logContent))
                    logContent = e.ToString();
                else
                    logContent += $" [Exception]: {e}";
            }

            string logMessage = $"{now.ToString(Properties.Resources.strLogFormat, CultureInfo.InvariantCulture)} [{type}]: {logContent}\n";
            System.Diagnostics.Debug.WriteLine(logMessage);

            // Append message to file
            try
            {
                // If log is too long, clear it
                var fi = new System.IO.FileInfo(LOG_PATH);
                if (fi.Exists && fi.Length > 500 * 1024 * 1024) // 500 MB
                    fi.Delete();

                if (!System.IO.File.Exists(LOG_PATH))
                    System.IO.File.WriteAllText(LOG_PATH, logMessage);
                else
                    System.IO.File.AppendAllText(LOG_PATH, logMessage);
            }
            catch
            {
                // If this fails ... UF ...
            }
        }

        [Flags]
        public enum LogType
        {
            Debug = 0x01,
            Info = 0x02,
            Warning = 0x04,
            Error = 0x08
        }
    }
}
