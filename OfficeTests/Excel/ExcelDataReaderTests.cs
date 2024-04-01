using me.fengyj.CommonLib.Office.Excel;


namespace me.fengyj.CommonLib.OfficeTests.Excel {
    [TestClass]
    public class ExcelDataReaderTests {

        [TestMethod]
        public void TestManualInputedNumber() {

            using (var reader = new ExcelDataReader<NumberTable>(Path.Combine("TestData", "ManualCreated.xlsx"))) {

                var cfg = new ExcelDataReader<NumberTable>.Config(
                    1,
                    new ExcelDataReader<NumberTable>.DataArea(2, 1, 8, 3),
                    c => new NumberTable(),
                    new List<ExcelDataReader<NumberTable>.DataDeserializer>() {
                        new ExcelDataReader<NumberTable>.TextDeserializer((o, v) => o.Comment = v, colIdx: 1),
                        new ExcelDataReader<NumberTable>.NumberDeserializer<double>(
                            (o, v) => o.ValueToCheck = v,
                            colIdx: 2),
                        new ExcelDataReader<NumberTable>.NumberDeserializer<double>(
                            (o, v) => o.ExpectedValue = v,
                            colIdx: 3)
                    });

                var data = reader.Read(cfg).ToList();
                Assert.AreEqual(6, data.Count);
                foreach (var item in data) {
                    Assert.AreEqual(item.ExpectedValue, item.ValueToCheck);
                }
            }
        }

        [TestMethod]
        public void TestManualInputedDate() {

            using (var reader = new ExcelDataReader<DateTable>(Path.Combine("TestData", "ManualCreated.xlsx"))) {

                var cfg = new ExcelDataReader<DateTable>.Config(
                    2,
                    new ExcelDataReader<DateTable>.DataArea(2, 1, 8, 3),
                    c => new DateTable(),
                    new List<ExcelDataReader<DateTable>.DataDeserializer>() {
                        new ExcelDataReader<DateTable>.TextDeserializer((o, v) => o.Comment = v, colIdx: 1),
                        new ExcelDataReader<DateTable>.DateTimeDeserializer(
                            (o, v) => o.ValueToCheck = v,
                            new string[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "MM/dd/yyyy", "dd/MM/yyyy", "MM/dd/yyyy h:mm tt"},
                            colIdx: 2),
                        new ExcelDataReader<DateTable>.DateTimeDeserializer(
                            (o, v) => o.ExpectedValue = v,
                            new string[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss"},
                            colIdx: 3)
                    });

                var data = reader.Read(cfg).ToList();
                Assert.AreEqual(7, data.Count);
                foreach (var item in data) {
                    Assert.AreEqual(item.ExpectedValue, item.ValueToCheck);
                }
            }
        }

        [TestMethod]
        public void TestManualCreatedExcel() {

            using (var reader = new ExcelDataReader<List<object?>>(Path.Combine("TestData", "ManualCreated.xlsx"))) {

                var cfg = new ExcelDataReader<List<object?>>.Config(
                    3,
                    new ExcelDataReader<List<object?>>.DataArea(1, 1),
                    c => new List<object?>(),
                    null,
                    (rec, c, v) => { rec.Add(v); });

                var data = reader.Read(cfg).ToList();
                Assert.AreEqual(10, data.Count);
            }
        }

        public class DateTable {
            public string? Comment { get; set; }
            public DateTime? ValueToCheck { get; set; }
            public DateTime? ExpectedValue { get; set; }
        }

        public class NumberTable {
            public string? Comment { get; set; }
            public double? ValueToCheck { get; set; }
            public double? ExpectedValue { get; set; }
        }
    }
}
