using me.fengyj.CommonLib.Office.Excel;


namespace me.fengyj.CommonLib.OfficeTests.Excel {
    [TestClass]
    public class ExcelDataReaderTests {

        [TestMethod]
        public void TestManualInputedNumber() {

            using var reader = new ExcelDataReader<NumberTable>(Path.Combine("TestData", "ManualCreated.xlsx"));
            var cfg = new ExcelDataReader<NumberTable>.Config(
                1,
                new ExcelDataReader<NumberTable>.DataArea(2, 1, 10, 3),
                c => new NumberTable(),
                [
                    new ExcelDataReader<NumberTable>.TextDeserializer((o, v) => o.Comment = v, colIdx: 1),
                        new ExcelDataReader<NumberTable>.NumberDeserializer<double>(
                            (o, v) => o.ValueToCheck_Double = v,
                            colIdx: 2),
                        new ExcelDataReader<NumberTable>.NumberDeserializer<double>(
                            (o, v) => o.ExpectedValue_Double = v,
                            colIdx: 3),
                        new ExcelDataReader<NumberTable>.NumberDeserializer<decimal>(
                            (o, v) => o.ValueToCheck_Decimal = v,
                            colIdx: 2),
                        new ExcelDataReader<NumberTable>.NumberDeserializer<decimal>(
                            (o, v) => o.ExpectedValue_Decimal = v,
                            colIdx: 3),
                        new ExcelDataReader<NumberTable>.TextDeserializer(
                            (o, v) => o.ValueToCheck_String = v,
                            colIdx: 2),
                        new ExcelDataReader<NumberTable>.TextDeserializer(
                            (o, v) => o.ExpectedValue_String = v,
                            colIdx: 3)
                ]);

            var data = reader.Read(cfg).Select(i => i.Data).ToList();
            Assert.AreEqual(9, data.Count);
            foreach (var item in data) {
                Assert.IsNotNull(item);
                Assert.IsTrue(
                    item.ExpectedValue_Double.HasValue && item.ExpectedValue_Double.Value.CompareTo(item.ValueToCheck_Double) < 0.000000000001,
                    $"Expected: {item.ExpectedValue_Double}, Actual: {item.ValueToCheck_Double}");
                Assert.AreEqual(item.ExpectedValue_Decimal, item.ValueToCheck_Decimal);
                Assert.AreEqual(item.ExpectedValue_String, item.ValueToCheck_String);
            }
        }

        [TestMethod]
        public void TestManualInputedDate() {

            using var reader = new ExcelDataReader<DateTable>(Path.Combine("TestData", "ManualCreated.xlsx"));
            var cfg = new ExcelDataReader<DateTable>.Config(
                2,
                new ExcelDataReader<DateTable>.DataArea(2, 1, 8, 3),
                c => new DateTable(),
                [
                    new ExcelDataReader<DateTable>.TextDeserializer((o, v) => o.Comment = v, colIdx: 1),
                        new ExcelDataReader<DateTable>.DateTimeDeserializer(
                            (o, v) => o.ValueToCheck = v,
                            ["yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "MM/dd/yyyy", "dd/MM/yyyy", "MM/dd/yyyy h:mm tt"],
                            colIdx: 2),
                        new ExcelDataReader<DateTable>.DateTimeDeserializer(
                            (o, v) => o.ExpectedValue = v,
                            ["yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss"],
                            colIdx: 3)
                ]);

            var data = reader.Read(cfg).Select(i => i.Data).ToList();
            Assert.AreEqual(7, data.Count);
            foreach (var item in data) {
                Assert.IsNotNull(item);
                Assert.AreEqual(item.ExpectedValue, item.ValueToCheck);
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

                var data = reader.Read(cfg).Select(i => i.Data).ToList();
                Assert.AreEqual(8, data.Count);
            }
        }

        public class DateTable {
            public string? Comment { get; set; }
            public DateTime? ValueToCheck { get; set; }
            public DateTime? ExpectedValue { get; set; }
        }

        public class NumberTable {
            public string? Comment { get; set; }
            public double? ValueToCheck_Double { get; set; }
            public double? ExpectedValue_Double { get; set; }
            public decimal? ValueToCheck_Decimal { get; set; }
            public decimal? ExpectedValue_Decimal { get; set; }
            public string? ValueToCheck_String { get; set; }
            public string? ExpectedValue_String { get; set; }
        }
    }
}
