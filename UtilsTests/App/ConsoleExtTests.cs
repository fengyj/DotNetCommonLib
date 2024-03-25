using System.Text.RegularExpressions;

using me.fengyj.CommonLib.Utils.App;

namespace UtilsTests.App {
    [TestClass]
    public class ConsoleExtTests {
        [TestMethod]
        public void Test() {

            foreach (var c in Enum.GetValues<ConsoleColor>())
                ConsoleExt.Write(c, $"{c} Text; ");
            Console.WriteLine();

            ConsoleExt.Write(ConsoleColor.Gray, "format text in {0}", ConsoleColor.Gray);
            Console.WriteLine();

            ConsoleExt.Write("format text in {0}, not {1}", ConsoleColor.Yellow, ConsoleColor.Yellow, ConsoleColor.Cyan);
            Console.WriteLine();

            ConsoleExt.WriteLine("format text in {0}, {1}, {2}", [ConsoleColor.Red, ConsoleColor.Green], null, ConsoleColor.Red, ConsoleColor.Green, ConsoleColor.Red);

            ConsoleExt.WriteLine("This is a message, and the word 'message' is marked as red.", new Regex("'?(?<word>message)'?"), [ConsoleColor.Blue, ConsoleColor.Red], null, ["word"]);
        }
    }
}
