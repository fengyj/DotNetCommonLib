using me.fengyj.CommonLib.Utils.Expressions;

namespace UtilsTests.Expressions {

    [TestClass]
    public class ExpressionUtilsTests {

        [TestMethod]
        public void Test_GetPropertyPath() {
            var path = ExpressionUtils.GetPropertyPath(propertyExp: (ClassA a) => a.B == null ? null : a.B.Date);
        }
    }

    public class ClassA {
        public string? Name { get; set; }
        public ClassB? B { get; set; }
    }

    public class ClassB {
        public DateTime? Date { get; set; }
    }
}
