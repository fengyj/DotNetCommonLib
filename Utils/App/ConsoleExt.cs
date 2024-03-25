using System.Collections.Concurrent;
using System.Drawing;
using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Utils.App {
    public static partial class ConsoleExt {

        private static readonly ConcurrentQueue<Action> writeActions = [];
        private static readonly SemaphoreSlim writeSemaphore = new(0);
        private static readonly CancellationTokenSource cancelTokenSource = new();
        private static readonly Regex formatStrSplit = FormatStringSplitRegex();
        private static volatile Task? writeAsyncTask = null;

        public static bool IsThreadSafeRequired { get; set; } = false;

        public static void Write(string str, ConsoleColor? foreColor = null, ConsoleColor? backColor = null) {

            WriteToConsole(() => Console.Write(str), foreColor, backColor);
        }

        public static void Write(ConsoleColor foreColor, string format, params object[] args) {

            WriteToConsole(() => Console.Write(format, args), foreColor: foreColor);
        }

        public static void Write(string format, ConsoleColor? foreColor, params object[] args) {

            Write(format, foreColors: foreColor == null ? null : [foreColor.Value], backColors: null, args);
        }

        /// <summary>
        /// Print message by giving format, and the args. The args will be highlighted with the color schemes.
        /// </summary>
        /// <example>
        /// <code>
        /// // the below code will output: format text in Red, Green, Red. And the "Red", "Green" in the messsage will be highlighted in red and green color.
        /// ConsoleExt.WriteLine("format text in {0}, {1}, {2}", [ConsoleColor.Red, ConsoleColor.Green], null, ConsoleColor.Red, ConsoleColor.Green, ConsoleColor.Red);
        /// </code>
        /// </example>
        /// <param name="format"></param>
        /// <param name="foreColors"></param>
        /// <param name="backColors"></param>
        /// <param name="args"></param>
        public static void Write(string format, ConsoleColor[]? foreColors, ConsoleColor[]? backColors, params object[] args) {

            WriteToConsole(GetWriteAction(format, foreColors, backColors, args));
        }

        public static Action GetWriteAction(string format, ConsoleColor[]? foreColors, ConsoleColor[]? backColors, params object[] args) {

            if (string.IsNullOrWhiteSpace(format)) return () => { };

            if (args == null || args.Length == 0) {

                return () => WriteToConsole(
                    () => Console.Write(format),
                    foreColor: foreColors == null || foreColors.Length == 0 ? null : foreColors[0],
                    backColor: backColors == null || backColors.Length == 0 ? null : backColors[0],
                    concurrentWrite: false);
            }

            if ((foreColors == null || foreColors.Length == 0) && (backColors == null || backColors.Length == 0)) {

                return () => Console.Write(format, args);
            }

            if (foreColors != null && foreColors.Length == 0) foreColors = null;
            if (backColors != null && backColors.Length == 0) backColors = null;

            var matches = formatStrSplit.Matches(format);

            var idx = 0;
            var idxOfForeColor = 0;
            var idxOfBackColor = 0;

            var actions = new List<Action>(); // put all actions in a list, then as a whole action to the Write function, for avoiding the concurrency issue.

            foreach (Match match in matches) {

                var lastPos = idx;
                var toPos = match.Index;
                actions.Add(() => Console.Write(format[lastPos..toPos]));

                idx = match.Index + match.Length;

                if (int.TryParse(match.Groups["id"]?.Value, out var i)) {

                    var s = args[i]?.ToString();
                    if (s == null) continue;

                    var foreColor = foreColors?[idxOfForeColor++ % foreColors.Length];
                    var backColor = backColors?[idxOfBackColor++ % backColors.Length];

                    actions.Add(() => WriteToConsole(
                        () => Console.Write(s),
                        foreColor: foreColor,
                        backColor: backColor,
                        concurrentWrite: false));
                }
            }

            if (idx < format.Length - 1) {
                var lastPos = idx;
                actions.Add(() => Console.Write(format[lastPos..]));
            }

            return () => actions.ForEach(a => a());
        }

        /// <summary>
        /// Print the str, and the values matched highlightedRegex will be highlighted with the giving color schemes.
        /// </summary>
        /// <example>
        /// <code>
        /// The code below prints: This is a message, and the word 'message' is marked as red. The value "'" surround the "message" is printed in blue, the "message" is printed in red, and the rest parts are in white.
        /// ConsoleExt.WriteLine("This is a message, and the word 'message' is marked as red.", new Regex("'?(?<word>message)'?"), [ConsoleColor.White, ConsoleColor.Blue, ConsoleColor.Red], null, ["word"]);
        /// </code>
        /// </example>
        /// <param name="str"></param>
        /// <param name="highlightRegex"></param>
        /// <param name="foreColors">the first one is for whole message, the first one (if have) is for the text matched the regex,
        /// the next one (if have, or use the color for the matched text), used for the named captured text.</param>
        /// <param name="backColors"></param>
        /// <param name="names">names in the named capturing group</param>
        public static void Write(string str, Regex highlightRegex, ConsoleColor[]? foreColors, ConsoleColor[]? backColors, params string[] names) {

            WriteToConsole(GetWriteAction(str, highlightRegex, foreColors, backColors, names));
        }

        /// <summary>
        /// Get an action object for printing the str, and the values matched highlightedRegex will be highlighted with the giving color schemes.
        /// </summary>
        /// <example>
        /// <code>
        /// The code below prints: This is a message, and the word 'message' is marked as red. The value "'" surround the "message" is printed in blue, the "message" is printed in red, and the rest parts are in white.
        /// ConsoleExt.WriteLine("This is a message, and the word 'message' is marked as red.", new Regex("'?(?<word>message)'?"), [ConsoleColor.White, ConsoleColor.Blue, ConsoleColor.Red], null, ["word"]);
        /// </code>
        /// </example>
        /// <param name="str"></param>
        /// <param name="highlightRegex"></param>
        /// <param name="foreColors">the first one is for whole message, the first one (if have) is for the text matched the regex,
        /// the next one (if have, or use the color for the matched text), used for the named captured text.</param>
        /// <param name="backColors"></param>
        /// <param name="names">names in the named capturing group</param>
        public static Action GetWriteAction(string str, Regex highlightRegex, ConsoleColor[]? foreColors, ConsoleColor[]? backColors, params string[] names) {

            if (string.IsNullOrWhiteSpace(str)) return () => { };

            if ((foreColors == null || foreColors.Length == 0) && (backColors == null || backColors.Length == 0)) {

                return () => Console.Write(str);
            }

            var actions = new List<Action>(); // put all actions in a list, then as a whole action to the Write function, for avoiding the concurrency issue.

            if (foreColors != null && foreColors.Length == 0) foreColors = null;
            if (backColors != null && backColors.Length == 0) backColors = null;

            var matches = highlightRegex.Matches(str);

            var idx = 0;
            var nameOfIndices = names?.Select((i, idx) => Tuple.Create(i, idx)).ToDictionary(i => i.Item1, i => i.Item2);

            var foreColorOfMatch = foreColors?[1 % foreColors.Length];
            var backColorOfMatch = backColors?[1 % backColors.Length];

            foreach (Match match in matches) {

                var lastPos = idx;
                var toPos = match.Index;
                actions.Add(() => WriteToConsole(() => Console.Write(str[lastPos..toPos]), foreColor: foreColors?[0], backColor: backColors?[0], concurrentWrite: false));

                idx = match.Index;

                for (var g = 1; g < match.Groups.Count; g++) {

                    var group = match.Groups[g];

                    var posBeforeGroup = idx;
                    var posOfGroupStart = group.Index;

                    actions.Add(() => WriteToConsole(() => Console.Write(str[posBeforeGroup..posOfGroupStart]), foreColor: foreColorOfMatch, backColor: backColorOfMatch, concurrentWrite: false));

                    idx = group.Index;

                    if (nameOfIndices?.TryGetValue(group.Name, out var idxOfName) ?? false) {

                        var posStart = idx;
                        var posEnd = idx + group.Length;
                        var foreColor = foreColors?[foreColors.Length > 2 ? ((idxOfName) % (foreColors.Length - 2) + 2) : 1 % foreColors.Length];
                        var backColor = backColors?[backColors.Length > 2 ? ((idxOfName) % (backColors.Length - 2) + 2) : 1 % backColors.Length];

                        actions.Add(() => WriteToConsole(() => Console.Write(str[posStart..posEnd]), foreColor: foreColor, backColor: backColor, concurrentWrite: false));

                        idx = group.Index + group.Length;
                    }
                }

                var posAfterGroup = idx;
                var posEndOfMatch = match.Index + match.Length;
                actions.Add(() => WriteToConsole(() => Console.Write(str[posAfterGroup..posEndOfMatch]), foreColor: foreColorOfMatch, backColor: backColorOfMatch, concurrentWrite: false));

                idx = match.Index + match.Length;
            }
            if (idx < str.Length - 1) {
                var lastPos = idx;
                actions.Add(() => WriteToConsole(() => Console.Write(str[lastPos..]), foreColor: foreColors?[0], backColor: backColors?[0], concurrentWrite: false));
            }

            return () => actions.ForEach(a => a());
        }

        /// <summary>
        /// Used for executing mutiple wirte functions as one action to avoid concurrency issue.
        /// </summary>
        /// <param name="actions">only supports the actions returned by GetWriteAction functions</param>
        public static void Write(params Action[] actions) {

            if (actions == null) return;

            WriteToConsole(() => {
                foreach (var item in actions) {
                    item();
                }
            });
        }

        public static void WriteLine(string str, ConsoleColor? foreColor, ConsoleColor? backColor = null) {

            WriteToConsole(() => Console.WriteLine(str), foreColor, backColor);
        }

        public static void WriteLine(ConsoleColor foreColor, string format, params object[] args) {
            WriteToConsole(() => Console.WriteLine(format, args), foreColor: foreColor);
        }

        public static void WriteLine(string format, ConsoleColor? foreColor, params object[] args) {

            WriteLine(format, foreColors: foreColor == null ? null : [foreColor.Value], backColors: null, args);
        }

        public static void WriteLine(string format, ConsoleColor[]? foreColors, ConsoleColor[]? backColors, params object[] args) {

            WriteLine(GetWriteAction(format, foreColors, backColors, args));
        }

        public static void WriteLine(string str, Regex highlightRegex, ConsoleColor[]? foreColors, ConsoleColor[]? backColors, params string[] names) {

            WriteLine(GetWriteAction(str, highlightRegex, foreColors, backColors, names));
        }

        /// <summary>
        /// Used for executing mutiple wirte functions as one action to avoid concurrency issue.
        /// </summary>
        /// <param name="actions">only supports the actions returned by GetWriteAction functions</param>
        public static void WriteLine(params Action[] actions) {

            if (actions == null) return;

            WriteToConsole(() => {

                foreach (var item in actions) {
                    item();
                }
                Console.WriteLine();
            });
        }

        private static void WriteToConsole(Action action, ConsoleColor? foreColor = null, ConsoleColor? backColor = null, bool? concurrentWrite = null) {

            if (concurrentWrite ?? IsThreadSafeRequired) {

                var writeAction = () => WriteToConsole(action, foreColor, backColor, false);
                writeActions.Enqueue(writeAction);
                if (writeAsyncTask == null) {
                    lock (typeof(ConsoleExt)) {
                        if (writeAsyncTask == null) {
                            AppDomain.CurrentDomain.ProcessExit += (sender, e) => cancelTokenSource.Cancel();
                            writeAsyncTask = WriteAsync();
                            writeAsyncTask.Start();
                        }
                    }
                }
                writeSemaphore.Release();
            }
            else {
                var origForeColor = Console.ForegroundColor;
                var origBackColor = Console.BackgroundColor;

                if (foreColor != null)
                    Console.ForegroundColor = foreColor.Value;
                if (backColor != null)
                    Console.BackgroundColor = backColor.Value;

                action();

                if (foreColor != null)
                    Console.ForegroundColor = origForeColor;
                if (backColor != null)
                    Console.BackgroundColor = origBackColor;
            }
        }

        private static async Task WriteAsync() {

            var token = cancelTokenSource.Token;
            while (!token.IsCancellationRequested) {
                await writeSemaphore.WaitAsync(token);
                if (writeActions.TryDequeue(out var writeAction))
                    writeAction();
            }
        }

        public static ConsoleColor ToNearestConsoleColor(Color color) {

            var closestConsoleColor = ConsoleColor.Black;
            var delta = double.MaxValue;

            foreach (ConsoleColor consoleColor in Enum.GetValues(typeof(ConsoleColor))) {

                var consoleColorName = Enum.GetName<ConsoleColor>(consoleColor);
                consoleColorName = string.Equals(consoleColorName, nameof(ConsoleColor.DarkYellow), StringComparison.Ordinal) ? nameof(Color.Orange) : consoleColorName;
                if (consoleColorName == null) continue;

                var rgbColor = Color.FromName(consoleColorName);
                var sum = Math.Pow(rgbColor.R - color.R, 2.0) + Math.Pow(rgbColor.G - color.G, 2.0) + Math.Pow(rgbColor.B - color.B, 2.0);

                var epsilon = 0.001;
                if (sum < epsilon) {
                    return consoleColor;
                }

                if (sum < delta) {
                    delta = sum;
                    closestConsoleColor = consoleColor;
                }
            }

            return closestConsoleColor;
        }

        [GeneratedRegex("\\{(?<id>\\d+)\\}", RegexOptions.Compiled)]
        private static partial Regex FormatStringSplitRegex();
    }
}
