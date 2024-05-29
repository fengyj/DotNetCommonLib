using me.fengyj.CommonLib.Office.Excel;
using me.fengyj.CommonLib.Utils;

namespace me.fengyj.CommonLib.OfficeTests.Excel {
    [TestClass]
    public class ExcelUtilTests {

        [TestMethod]
        public void Test_Read() {

            var tmpFile = IOUtil.GetTempFile(fileExt: ".xlsx");

            try {

                var builder = new ExcelBuilder();

                builder.AppendSheet("sheet1")
                    .AddRows(Enumerable.Repeat(new object[] { 1, "str", DateTime.Today, 0.2 }, 10));

                builder.AppendSheet("sheet2")
                    .AddRow(100, rowOffset: 2, colOffset: 3);

                builder.AppendSheet("sheet3")
                    .AddTable(Enumerable.Repeat(new object[] { 123456m, true, DateTime.Today }, 2), new TableConfig<object[]>(new List<ITableColumnConfig<object[]>>() {
                        new TableColumnConfig<object[], object>("col1", dataGetter: i => i[0]),
                        new TableColumnConfig<object[], object>("col2", dataGetter: i => i[1]),
                        new TableColumnConfig<object[], object>("col3", dataGetter: i => i[2], style: CellStyle.DateTime_Default.With(numberingStyle: NumberingStyle.DateTime_UK)),
                        new TableColumnConfig<object[], object>("col4", dataGetter: i => i[2], style: CellStyle.DateTime_Default.With(numberingStyle: NumberingStyle.Date_US))
                    }));


                builder.BuildTo(tmpFile.FullName);

                var data = ExcelUtil.Read(tmpFile.FullName, 1, 1, 1, 10, 4).Select(i => i.Data); // sheet1
                var count = 0;
                foreach (var row in data) {
                    count++;
                    Assert.IsNotNull(row);
                    Assert.AreEqual("1", row[0]);
                    Assert.AreEqual("str", row[1]);
                    Assert.AreEqual(DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss.fff"), row[2]);
                    Assert.AreEqual("0.2", row[3]);
                }
                Assert.AreEqual(10, count);

                data = ExcelUtil.Read(tmpFile.FullName, 2, 2, 3).Select(i => i.Data); // sheet2
                Assert.AreEqual(1, data.Sum(i => i?.Count ?? 0));

                var reused = new List<string?>();
                data = ExcelUtil.Read(tmpFile.FullName, 3, 1, 1, reused: reused).Select(i => i.Data);
                count = 0;
                foreach (var row in data) {
                    count++;
                    Assert.IsNotNull(row);
                    if (count == 1) {
                        Assert.AreEqual("col1", row[0]);
                        Assert.AreEqual("col2", row[1]);
                        Assert.AreEqual("col3", row[2]);
                        Assert.AreEqual("col4", row[3]);
                    }
                    else {
                        Assert.AreEqual("123456", row[0]);
                        Assert.AreEqual("true", row[1]);
                        Assert.AreEqual(DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss.fff"), row[2]);
                        Assert.AreEqual(DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss.fff"), row[3]);
                    }
                }
                Assert.AreEqual(3, count);

            }
            finally {
                IOUtil.DeleteFile(tmpFile);
            }
        }
    }
}
