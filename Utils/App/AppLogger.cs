using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Utils.App {

    /// <summary>
    /// This is just used by this library. By default, it just prints the loggers to console. 
    /// If needs to write them to somewhere else, please set the log functions.
    /// </summary>
    public class AppLogger {

        private static readonly ConsoleColor[] Color_Scheme_DEBUG = [ConsoleColor.DarkMagenta, ConsoleColor.Gray, ConsoleColor.White];
        private static readonly ConsoleColor[] Color_Scheme_INFO = [ConsoleColor.DarkMagenta, ConsoleColor.Green, ConsoleColor.White];
        private static readonly ConsoleColor[] Color_Scheme_WARN = [ConsoleColor.DarkMagenta, ConsoleColor.Yellow, ConsoleColor.White, ConsoleColor.Gray];
        private static readonly ConsoleColor[] Color_Scheme_ERROR = [ConsoleColor.DarkMagenta, ConsoleColor.Red, ConsoleColor.White, ConsoleColor.Gray];
        private static readonly ConsoleColor[] Color_Scheme_Highlight = [ConsoleColor.White, ConsoleColor.Cyan];

        private const string Log_Timestamp_Format = "yyyy-MM-dd HH:mm:ss.fff";

        public static Action<string, object[]?>? LogDebugFunc { get; set; }
        public static Action<string, object[]?>? LogInfoFunc { get; set; }
        public static Action<string, Exception?, object[]?>? LogWarnFunc { get; set; }
        public static Action<string, Exception?, object[]?>? LogErrorFunc { get; set; }

        public static void LogDebug(string msg, params object[] args) {

            if (LogDebugFunc != null)
                LogDebugFunc(msg, args);
            else
                Log("DEBUG", Color_Scheme_DEBUG, msg, null, args);
        }

        public static void LogInfo(string msg, params object[] args) {

            if (LogInfoFunc != null)
                LogInfoFunc(msg, args);
            else
                Log("INFO", Color_Scheme_INFO, msg, null, args);
        }

        public static void LogWarn(string msg, Exception? ex = null, params object[] args) {

            if (LogWarnFunc != null)
                LogWarnFunc(msg, ex, args);
            else
                Log("WARN", Color_Scheme_WARN, msg, ex, args);
        }

        public static void LogError(string msg, Exception? ex = null, params object[] args) {

            if (LogErrorFunc != null)
                LogErrorFunc(msg, ex, args);
            else
                Log("ERROR", Color_Scheme_ERROR, msg, ex, args);
        }

        private static void Log(string logLevel, ConsoleColor[] colorScheme, string msg, Exception? ex, params object[] args) {

            if (ex == null && (args == null || args.Length == 0)) {
                ConsoleExt.WriteLine("{0} [{1}] {2}", colorScheme, null, GetNow(), logLevel.PadRight(5), msg);
                return;
            }
            if (args != null && args.Where(a => !string.IsNullOrEmpty(a?.ToString())).Any()) {

                var argsInStr = args.Where(a => a != null).Select(a => a.ToString()).Select(selector: s => s == null ? null : Regex.Escape(s)).ToArray();
                var regex = new Regex($"({string.Join('|', argsInStr)})");

                ConsoleExt.WriteLine(
                    ConsoleExt.GetWriteAction("{0} [{1}] ", colorScheme, null, GetNow(), logLevel.PadRight(5)),
                    ConsoleExt.GetWriteAction(string.Format(msg, args), regex, foreColors: Color_Scheme_Highlight, backColors: null),
                    () => { if (ex != null) ConsoleExt.WriteLine(ex.ToString(), foreColor: colorScheme[colorScheme.Length - 1]); else Console.WriteLine(); });
            }
            else if (ex != null) {
                ConsoleExt.WriteLine("{0} [{1}] {2} {3}", colorScheme, null, GetNow(), logLevel.PadRight(5), msg, ex.ToString());
            }
        }

        private static string GetNow() {
            return DateTime.Now.ToString(Log_Timestamp_Format);
        }
    }
}
