using me.fengyj.CommonLib.Utils.Data;

namespace UtilsTests.Data {

    [TestClass]
    public class DataTableUtilTests {

        [TestMethod]
        public void DataTableUtilTest() {

            var table = DataTableUtil.CreateTable([["cell_1", 1], [null, null], [1, 2]], ["Column_A", "Column_B"], "table");

            Assert.AreEqual("table", table.TableName);
            Assert.AreEqual("Column_A", table.Columns[0].ColumnName);
            Assert.AreEqual("Column_B", table.Columns[1].ColumnName);
            Assert.AreEqual(typeof(string), table.Columns[0].DataType);
            Assert.AreEqual(typeof(int), table.Columns[1].DataType);
            Assert.AreEqual(3, table.Rows.Count);
            Assert.AreEqual("cell_1", table.Rows[0][0]);
            Assert.AreEqual(1, table.Rows[0][1]);
            Assert.AreEqual(DBNull.Value, table.Rows[1][0]);
            Assert.AreEqual(DBNull.Value, table.Rows[1][1]);
            Assert.AreEqual("1", table.Rows[2][0]);
            Assert.AreEqual(2, table.Rows[2][1]);

            table = DataTableUtil.CreateTable(null, ["a", "b"]);

            Assert.AreEqual(2, table.Columns.Count);
            Assert.AreEqual(0, table.Rows.Count);

            try {
                table = DataTableUtil.CreateTable(null);
                Assert.Fail("The code above should throw exception.");
            }
            catch (ArgumentException) { }
        }
    }
}
