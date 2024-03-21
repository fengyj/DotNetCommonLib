using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using me.fengyj.CommonLib.Office.Excel;
using me.fengyj.CommonLib.Utils;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace me.fengyj.CommonLib.OfficeTests.Excel {

    [TestClass]
    public class ExcelBuilderTests {

        [TestMethod]
        public void BuilderTest() {

            List<SampleData> data = [
                    new() {
                        StringValue = "Item 1",
                        IntValue = -1_000_000,
                        ByteValue = 2,
                        UShortValue = 4_235,
                        FloatValue = 0.234f,
                        DoubleValue = 34335423.03445,
                        DecimalValue = -3234.345323m,
                        DateTimeValue = DateTime.Now,
                        BooleanValue = false,
                        LongValue = -335326326353,
                        ULongValue = 3675663534345,
                        TimeSpanValue = TimeSpan.FromSeconds(3345423)
                    },
                    new() {
                        StringValue = "Item 2",
                        IntValue = -1_123_000_000,
                        ByteValue = 25,
                        UShortValue = 4_235,
                        FloatValue = 3540.234f,
                        DoubleValue = -34235423.0353445,
                        DecimalValue = -3234.345323m,
                        DateTimeValue = DateTime.Today,
                        BooleanValue = true,
                        LongValue = -3353263353,
                        ULongValue = 367563534345,
                        TimeSpanValue = null,
                    },
                    new(),
                    new() {
                        StringValue = "Item 4",
                        IntValue = null,
                        ByteValue = null,
                        UShortValue = null,
                        FloatValue = null,
                        DoubleValue = null,
                        DecimalValue = null,
                        DateTimeValue = null,
                        BooleanValue = null,
                        LongValue = null,
                        ULongValue = null,
                        TimeSpanValue = TimeSpan.FromSeconds(34561)
                    }
                ];
            int id = 1;
            var tblCfg = new TableConfig<SampleData>([
                    new TableColumnConfig<SampleData, string>("ID Column", dataGetter: i => id++.ToString()),
                    new TableColumnConfig<SampleData, string?>("String Column", dataGetter: i => i.StringValue),
                    new TableColumnConfig<SampleData, int?>("Int Column", dataGetter: i => i.IntValue),
                    new TableColumnConfig<SampleData, byte?>("Byte Column", dataGetter: i => i.ByteValue),
                    new TableColumnConfig<SampleData, ushort?>("UShort Column", dataGetter: i => i.UShortValue),
                    new TableColumnConfig<SampleData, float?>("Float Column", dataGetter: i => i.FloatValue),
                    new TableColumnConfig<SampleData, double?>("Double Column", dataGetter: i => i.DoubleValue),
                    new TableColumnConfig<SampleData, decimal?>("Decimal Column", dataGetter: i => i.DecimalValue),
                    new TableColumnConfig<SampleData, DateTime?>("DateTime Column", dataGetter: i => i.DateTimeValue),
                    new TableColumnConfig<SampleData, bool?>("Bool Column", dataGetter: i => i.BooleanValue),
                    new TableColumnConfig<SampleData, TimeSpan?>("TimeSpan Column", dataGetter: i => i.TimeSpanValue),
                    new TableColumnConfig<SampleData, long?>("Long Column", dataGetter: i => i.LongValue),
                    new TableColumnConfig<SampleData, ulong?>("ULong Column", dataGetter: i => i.ULongValue)
                    ],
                tableName: "Sample Data");

            var builder = new ExcelBuilder();

            var sheetBuilder = builder.AppendSheet("Sampel Data Demo", columnWidths: new Dictionary<int, uint> { { 1, 10u } });

            sheetBuilder.AddRow("Demo for ExcelBuilder", style: CellStyle.H1, rowOffset: 2);
            sheetBuilder.AddRow("It supports following types:", CellStyle.Emphasize, rowOffset: 2);
            sheetBuilder.AddRows(["int", "string", "DateTime", "TimeSpan", "decimal", "bool", "etc."], colOffset: 2);

            sheetBuilder.AddRow("It also supports to define table.", style: CellStyle.Emphasize, rowOffset: 2);
            sheetBuilder.AddTable(data, tblCfg, rowOffset: 2);

            sheetBuilder.AddRow(["Here is a hyperlink: ", new HyperLinkValue("http://bing.com", "Bing") ], rowOffset: 2);

            sheetBuilder.AddRow("And multiple sheets...", rowOffset: 2, style: CellStyle.Warning);

            sheetBuilder = builder.AppendSheet("An empty sheet");

            sheetBuilder = builder.AppendSheet("Table Styles");

            sheetBuilder.AddRow("No header:");
            sheetBuilder.AddTable(
                [[(object)"abc", (object)123], ["def", 456], [null, 42533]],
                new TableConfig<object[]>([
                    new TableColumnConfig<object[], object>("Col_1", dataGetter: i => i[0]),
                    new TableColumnConfig<object[], object>("Col_2", dataGetter: i => i[1])],
                style: new TableStyle(showHeader: false)));

            sheetBuilder.AddRow("Different theme:", rowOffset: 2, style: CellStyle.Error);
            sheetBuilder.AddTable(
                [[(object)"abc", (object)123], ["def", 456], [null, 42533]],
                new TableConfig<object[]>([
                    new TableColumnConfig<object[], object>("Col_1", dataGetter: i => i[0]),
                    new TableColumnConfig<object[], object>("Col_2", dataGetter: i => i[1])], 
                style: new TableStyle(styleName: "TableStyleMedium3")));

            sheetBuilder.AddRow("Different date format:", rowOffset: 2, style: CellStyle.Quote);
            sheetBuilder.AddTable(
                [[DateTime.Now, DateTime.Today, DateTime.UtcNow]],
                new TableConfig<DateTime[]>([
                    new TableColumnConfig<DateTime[], DateTime>("Default", dataGetter: i => i[0], style: CellStyle.Cell_DateTime_Default),
                    new TableColumnConfig<DateTime[], DateTime>("UK", dataGetter: i => i[1], style: CellStyle.Cell_Date_Default.With(numberingStyle: NumberingStyle.Date_UK)),
                    new TableColumnConfig<DateTime[], DateTime>("US", dataGetter: i => i[2], style: CellStyle.Cell_Date_Default.With(numberingStyle: NumberingStyle.DateTime_US))]));

            sheetBuilder.AddRow("Different number format:", rowOffset: 2, style: CellStyle.Quote);
            sheetBuilder.AddTable(
                [[123456789, 1024.4201, 1000000.123m, 0.123]],
                new TableConfig<object[]>([
                    new TableColumnConfig<object[], object>("Integer", dataGetter: i => i[0], style: CellStyle.Cell_Integer_Default),
                    new TableColumnConfig<object[], object>("Integer 2", dataGetter: i => i[0], style: CellStyle.Cell_Integer_Default.With(numberingStyle: NumberingStyle.Integer_Thousands)),
                    new TableColumnConfig<object[], object>("Decimal", dataGetter: i => i[1], style: CellStyle.Cell_Decimal_Default),
                    new TableColumnConfig<object[], object>("Decimal with 2 digits", dataGetter: i => i[2], style: CellStyle.Cell_Decimal_Default.With(numberingStyle: NumberingStyle.Decimal_Thousands_2)),
                    new TableColumnConfig<object[], object>("Percentage", dataGetter: i => i[3], style: CellStyle.Cell_Decimal_Default.With(numberingStyle: NumberingStyle.Percent_1))]));

            sheetBuilder.AddRow(["Here is a hyperlink: ", new HyperLinkValue("http://google.com", "Google")], rowOffset: 2);

            var tmpFile = IOUtil.GetTempFile(fileExt: ".xlsx");
            try {
                builder.BuildTo(tmpFile.FullName);
            }
            finally {
                IOUtil.DeleteFile(tmpFile);
            }
        }

        public class SampleData {
            public string? StringValue { get; set; }
            public int? IntValue { get; set; }
            public byte? ByteValue { get; set; }
            public ushort? UShortValue { get; set; }
            public float? FloatValue { get; set; }
            public double? DoubleValue { get; set; }
            public decimal? DecimalValue { get; set; }
            public DateTime? DateTimeValue { get; set; }
            public bool? BooleanValue { get; set; }
            public long? LongValue { get; set; }
            public ulong? ULongValue { get; set; }
            public TimeSpan? TimeSpanValue { get; set; }
        }
    }
}
