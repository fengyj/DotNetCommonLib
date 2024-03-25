namespace me.fengyj.CommonLib.Utils.App {
    public class AppArguments<T> {

        public const string Arg_Help_Name = "h";

        public AppArguments(T value, string[] inputs, string appDescription = "", params Arg[] args) {

            this.AppDescription = appDescription;
            this.Value = value;

            if (args == null) return;
            foreach (var item in args)
                this.Args.TryAdd(item.Name, item);

            this.Parse(inputs);
        }

        public Dictionary<string, Arg> Args { get; private set; } = [];
        public string AppDescription { get; private set; }
        public T Value { get; private set; }
        public bool IsHelpInfoRequired { get; private set; } = false;

        private void Parse(string[] inputs) {

            if (inputs == null || inputs.Length == 0) return;

            foreach (var item in inputs) {

                var keyValue = item.Split('=');
                var key = keyValue[0].Trim().TrimStart('-');
                var val = keyValue.Length == 2 ? keyValue[1].Trim() : null;

                if (key == Arg_Help_Name) {

                    this.IsHelpInfoRequired = true;
                }
                else if (this.Args.TryGetValue(key, out var arg)) {

                    try {
                        arg.ValueParser(this.Value, val ?? arg.DefaultValue);
                    }
                    catch (Exception ex) {
                        throw new ArgumentException("Cannot parse the parameter value.", arg.Name, ex);
                    }
                }
                else {
                    AppLogger.LogWarn($"Unknow argument {key} was found in the arguments.");
                }
            }
        }

        public bool PrintHelpInfoIfRequested() {

            if (!this.IsHelpInfoRequired) return false;

            Console.WriteLine();
            if (!string.IsNullOrWhiteSpace(this.AppDescription)) {
                ConsoleExt.WriteLine(this.AppDescription, ConsoleColor.Gray);
                Console.WriteLine();
            }

            var maxArgNameLength = this.Args.Select(i => i.Key.Length).Max();
            var len = (int)Math.Ceiling((maxArgNameLength + 1) / 4.0) * 4;
            Func<string, string> argNameFormatter = name => name.PadRight(len, ' ');

            var colorScheme = new ConsoleColor[] { ConsoleColor.DarkYellow, ConsoleColor.White, ConsoleColor.Cyan };

            ConsoleExt.WriteLine("  -{0}: {1}", colorScheme, null, argNameFormatter(Arg_Help_Name), "Print help information.");
            foreach (var arg in this.Args.Values) {

                if (arg.DefaultValue == null)
                    ConsoleExt.WriteLine("  -{0}: {1}", colorScheme, null, argNameFormatter(arg.Name), arg.Description);
                else
                    ConsoleExt.WriteLine("  -{0}: {1} Default value: {2}", colorScheme, null, argNameFormatter(arg.Name), arg.Description, arg.DefaultValue); ;
            }
            Console.WriteLine();
            return true;
        }

        public void LogArguments() {

            var args = string.Join(", ", this.Args.Select(i => Tuple.Create(i.Key, i.Value.ValueStringer(this.Value)))
                .Where(i => i.Item2 != null).Select(i => $"{i.Item1}={i.Item2}"));

            AppLogger.LogInfo($"{AppDomain.CurrentDomain.FriendlyName} is started with the arguments: {{0}}", args);
        }

        public class Arg {

            public Arg(string name, Action<T, string?> valueParser, Func<T, string?> valueStringier, string desc = "", string? defaultValue = null) {
                this.Name = name;
                this.ValueParser = valueParser;
                this.ValueStringer = valueStringier;
                this.DefaultValue = defaultValue;
                this.Description = desc;
            }

            public string Name { get; private set; }
            public Action<T, string?> ValueParser { get; private set; }
            public Func<T, string?> ValueStringer { get; private set; }
            public bool Optional { get; private set; } = false;
            public string? DefaultValue { get; private set; }
            public string Description { get; private set; }
        }
    }
}