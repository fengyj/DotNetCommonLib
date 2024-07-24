using me.fengyj.CommonLib.Utils.Expressions;

namespace UtilsTests.Expressions {

    [TestClass]
    public class ExpressionUtilsTests {

        [TestMethod]
        public void Test_GetPropertyPath() {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            var path = ExpressionUtils.GetPropertyPath(propertyExp: (ClassA a) => a.B.Date);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        [TestMethod]
        public void Test_GetPropertyExp() {

            var exp = ExpressionUtils.GetPropertySetterExp<ClassA, string>(nameof(ClassA.Name));
            var func = exp.Compile();
            var a = new ClassA();
            func(a, "E");
            Assert.AreEqual("E", a.Name);
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
