// See https://aka.ms/new-console-template for more information
using System.Configuration;
using System.Text.RegularExpressions;

using me.fengyj.CommonLib.Utils.App;

using TestConsoleApp;

internal class Program {
    static void Main(string[] args) {

        var exeFileName = Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly()?.GetName().Name);

        var cfgFileMap = new System.Configuration.ExeConfigurationFileMap();
        var srcFile = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            AppEnvironment.Default.ConfigsFolder,
            $"{exeFileName}.dev.config");
        var tagFile = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            $"{exeFileName}.exe.config");

        File.Copy(srcFile, tagFile, true);

        var cfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        foreach (ConfigurationSection s in cfg.Sections) {
            ConfigurationManager.RefreshSection(s.SectionInformation.Name);
        }

        args = ["-date=2024-01-01", "-h"];
        var appArgs = new AppArguments<Args>(
            new Args(),
            args,
            appDescription: "This is a program for testing prupose.",
            new AppArguments<Args>.Arg(
                "date",
                (obj, str) => {
                    if (DateTime.TryParse(str, out var v)) obj.Date = v;
                },
                obj => obj.Date?.ToString("yyyy-MM-dd"),
                desc: "Date"));

        appArgs.LogArguments();
        appArgs.PrintHelpInfoIfRequested();

        Console.WriteLine("Hello, World!");


        foreach (var c in Enum.GetValues<ConsoleColor>())
            ConsoleExt.Write(c, $"{c} Text; ");
        Console.WriteLine();

        ConsoleExt.Write(ConsoleColor.Gray, "format text in {0}", ConsoleColor.Gray);
        Console.WriteLine();

        ConsoleExt.Write("format text in {0}, not {1}", ConsoleColor.Yellow, ConsoleColor.Yellow, ConsoleColor.Cyan);
        Console.WriteLine();

        ConsoleExt.WriteLine("format text in {0}, {1}, {2}", [ConsoleColor.Red, ConsoleColor.Green], null, ConsoleColor.Red, ConsoleColor.Green, ConsoleColor.Red);

        ConsoleExt.WriteLine("This is a message, and the word 'message' is marked as red.", new Regex("'?(?<word>message)'?"), [ConsoleColor.White, ConsoleColor.Blue, ConsoleColor.Red], null, ["word"]);
    }

    public class Args {
        public DateTime? Date { get; set; }
    }
}