using me.fengyj.CommonLib.Utils;

namespace UtilsTests {
    [TestClass]
    public class StringTemplateUtilTest {
        [TestMethod]
        public void Test_GetText() {

            var template = "The $a's value is {$a,2}, __b's value is {__b:yyyyMMdd}, and c5's is {c5,-6:N2}.";
            var str = StringTemplateUtil.Default.GetText(template, new Dictionary<string, object> { { "$a", "A" }, { "__b", DateTime.Today }, { "c5", 20.1234 } });
            var expected = $"The $a's value is  A, __b's value is {DateTime.Today.ToString("yyyyMMdd")}, and c5's is 20.12 .";

            Assert.AreEqual(expected, str);
        }
    }
}
