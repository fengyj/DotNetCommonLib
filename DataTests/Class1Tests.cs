using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Data.Tests {
    [TestClass()]
    public class Class1Tests {
        [TestMethod()]
        public void Test_method1() {

            var regex = new Regex("\\{(?<id>\\d+)\\}");

            var matches = regex.Matches("sfsf {1}, sfdsfsfs{2}, fsdfsdf{1}");
            Assert.IsNotNull(matches);
        }
    }
}