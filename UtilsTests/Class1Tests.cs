using Microsoft.VisualStudio.TestTools.UnitTesting;
using me.fengyj.CommonLib.Utils;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace me.fengyj.CommonLib.Utils.Tests {
    [TestClass()]
    public class Class1Tests {
        [TestMethod()]
        public void Test_test() {
            Assert.Fail();
        }


        [TestMethod]
        public void test() {

            int? v = null;
            Assert.IsFalse(v is int);
            v = (int?)1;
            Assert.IsTrue(v is int);
            object o = (object)v;
            Assert.IsTrue(v is int);

            dosomething(v, typeof(int));
            v = new Nullable<int>(1);
            dosomething(v, typeof(int));
            v = null;
            dosomething(v, typeof(object));

            Func<int, int?> func = x => (int?)x;

            Assert.AreEqual(typeof(Nullable<int>), func.GetType().GenericTypeArguments[1]);

            DateTime? dt = DateTime.Now;
            dosomething(dt, typeof(DateTime));

        }

        private void dosomething(object? o, Type expectedType) {

            if (o is null) return;

            var t = o.GetType();

            Assert.AreEqual(expectedType, t);
        }
    }
}