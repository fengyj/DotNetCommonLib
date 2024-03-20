using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using me.fengyj.CommonLib.Office.Excel;

namespace me.fengyj.CommonLib.OfficeTests.Excel {
    [TestClass()]
    public class Test2 {
        string[] Sheet1Headers = ["Transaction Date", "Transaction Code", "Transaction Description", "Transaction Amount"];
        string[][] Sheet1Data = [
            ["2022-08-15", "SF", "SERVICE FEE", "9808.40"],
                ["2022-09-15", "DS", "DEPOSIT", "17808.40"],
                ["2022-10-15", "DS", "SERVICE FEE", "9808.40"],
                ["2022-11-15", "DS", "SERVICE FEE", "1508.40"],
                ["2022-12-15", "DS", "SERVICE FEE", "13208.40"]
        ];

        string[] Sheet2Headers = ["Timestamp", "Transit Number", "Unit", "Amount"];
        string[][] Sheet2Data = [
            ["2023-01-26T15:10:32.133-05:00", "124", "MB", "300.40"],
                ["2023-01-26T15:13:59.309-05:00", "278", "Minutes", "17808.40"],
                ["2023-01-26T15:52:43.730-05:00", "3693", "Percentage", "7.47"],
                ["2023-01-26T15:31:43.867-05:00", "4893", "Minutes", "1.40"],
                ["2023-01-26T15:32:15.037-05:00", "536", "Percentage", "0.19"]
        ];

        [TestMethod()]
        public void Test_Method1() {

            using (SpreadsheetDocument excel = SpreadsheetDocument.Create("demo1.xlsx", SpreadsheetDocumentType.Workbook)) {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = excel.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart1 = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart1.Worksheet = new Worksheet();
                SheetData data1 = CreateShee1Data();
                Columns columns1 = AutoSizeCells(data1);
                worksheetPart1.Worksheet.Append(columns1);
                worksheetPart1.Worksheet.Append(data1);
                DefineTable(worksheetPart1, 1, (uint)Sheet1Data.Length + 1, 1, (uint)Sheet1Headers.Length, Sheet1Headers, hasColumnHeader: false);

                WorksheetPart worksheetPart2 = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart2.Worksheet = new Worksheet();
                SheetData data2 = CreateShee2Data();
                Columns columns2 = AutoSizeCells(data2);
                worksheetPart2.Worksheet.Append(columns2);
                worksheetPart2.Worksheet.Append(data2);
                DefineTable(worksheetPart2, 1, (uint)Sheet2Data.Length + 1, 1, (uint)Sheet2Headers.Length, Sheet2Headers);

                var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                SheetStyleBuilder.BuildTo(workbookpart);

                // Add Sheets to the Workbook.
                Sheets sheets = excel.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                // Append a new worksheet and associate it with the workbook.
                sheets.Append(new Sheet() {
                    Id = excel.WorkbookPart.GetIdOfPart(worksheetPart1),
                    SheetId = 1,
                    Name = "Sheet 1"
                });
                sheets.Append(new Sheet() {
                    Id = excel.WorkbookPart.GetIdOfPart(worksheetPart2),
                    SheetId = 2,
                    Name = "Sheet 2"
                });

                //Save & close
                workbookpart.Workbook.Save();
                excel.Dispose();
            }
        }
        public static StringValue GetCellReference(uint row, uint column) =>
        new StringValue($"{GetColumnName("", column)}{row}");

        static string GetColumnName(string prefix, uint column) =>
            column < 26 ? $"{prefix}{(char)(65 + column - 1)}" :
            GetColumnName(GetColumnName(prefix, (column - column % 26) / 26 - 1), column % 26);

        static string GetTableReference(uint rowMin, uint rowMax, uint colMin, uint colMax) {
            return $"{GetCellReference(rowMin, colMin)}:{GetCellReference(rowMax, colMax)}";
        }
        int tableSeq = 0;
        void DefineTable(WorksheetPart worksheetPart, uint rowMin, uint rowMax, uint colMin, uint colMax, string[] headers, bool hasTotalRow = false, bool hasColumnHeader =true) {
            
            tableSeq++;
            var tableId = $"rId{tableSeq}";
            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>(tableId);
            int tableNo = worksheetPart.TableDefinitionParts.Count();

            string reference = GetTableReference(rowMin, rowMax, colMin, colMax);

            Table table = new Table() { Id = (UInt32)tableSeq, Name = "Table" + tableSeq, DisplayName = "Table" + tableSeq, Reference = reference, TotalsRowShown = hasTotalRow };

            TableColumns tableColumns = new TableColumns() { Count = (UInt32)(colMax - colMin + 1) };
            for (int i = 0; i < (colMax - colMin + 1); i++) {
                tableColumns.Append(new TableColumn() { Id = (UInt32)(colMin + i), Name = headers[i] }); //changed i+1 -> colMin + i
                                                                                                         //Add cell values (shared string)
            }

            TableStyleInfo tableStyleInfo = new TableStyleInfo() { Name = "TableStyleMedium1", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };
            if(!hasColumnHeader) {

                table.HeaderRowCount = UInt32Value.FromUInt32(0);
            }
            else {

                AutoFilter autoFilter = new AutoFilter() { Reference = reference };
                table.Append(autoFilter);
            }
            table.Append(tableColumns);
            table.Append(tableStyleInfo);

            tableDefinitionPart.Table = table;

            TableParts tableParts = (TableParts)worksheetPart.Worksheet.ChildElements.Where(ce => ce is TableParts).FirstOrDefault(); // Add table parts only once
            if (tableParts is null) {
                tableParts = new TableParts();
                //tableParts.Count = (UInt32)0;
                worksheetPart.Worksheet.Append(tableParts);
            }

            //tableParts.Count += (UInt32)1;
            TablePart tablePart = new TablePart() { Id = tableId };
            
            tableParts.Append(tablePart);

            return;
        }

        void InsertTextCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.InlineString, InlineString = new InlineString() { Text = new Text(content) } }, cellIndex);
        }

        void InsertHeaderCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() {
                DataType = CellValues.InlineString,
                InlineString = new InlineString() { Text = new Text(content) },
                StyleIndex = Office.Excel.CellStyle.TableHeader.StyleId
            }, cellIndex);
        }

        void InsertTransitCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Number, CellValue = new CellValue(content)/*, StyleIndex = 7*/ }, cellIndex);
        }

        void InsertNumberCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Number, CellValue = new CellValue(content)/*, StyleIndex = 5*/ }, cellIndex);
        }

        void InsertDateCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Date, CellValue = new CellValue(content)/*, StyleIndex = 1*/ }, cellIndex);
        }

        void InsertDateTimeCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Date, CellValue = new CellValue(DateTime.Parse(content)), StyleIndex = Office.Excel.CellStyle.Cell_DateTime_Default.StyleId /*StyleIndex = 3*/ }, cellIndex);
        }

        SheetData CreateShee1Data() {
            SheetData data = new SheetData();
            int rowId = 0;
            Row row = new Row();
            for (int i = 0; i < Sheet1Headers.Length; i++) {
                InsertHeaderCell(row, Sheet1Headers[i], i);
            }
            data.InsertAt(row, rowId++);

            for (int i = 0; i < Sheet1Data.Length; i++) {
                row = new Row();
                InsertDateCell(row, Sheet1Data[i][0], 0);
                InsertTextCell(row, Sheet1Data[i][1], 1);
                InsertTextCell(row, Sheet1Data[i][2], 2);
                InsertNumberCell(row, Sheet1Data[i][3], 3);
                data.InsertAt(row, rowId++);
            }

            return data;
        }

        SheetData CreateShee2Data() {
            SheetData data = new SheetData();
            int rowId = 0;
            Row row = new Row();
            for (int i = 0; i < Sheet2Headers.Length; i++) {
                InsertHeaderCell(row, Sheet2Headers[i], i);
            }
            data.InsertAt(row, rowId++);

            for (int i = 0; i < Sheet2Data.Length; i++) {
                row = new Row();
                InsertDateTimeCell(row, Sheet2Data[i][0], 0);
                InsertTransitCell(row, Sheet2Data[i][1], 1);
                InsertTextCell(row, Sheet2Data[i][2], 2);
                InsertNumberCell(row, Sheet2Data[i][3], 3);
                data.InsertAt(row, rowId++);
            }

            return data;
        }

        Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData) {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
            foreach (var r in rows) {
                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++) {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? cell.InnerText : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex)) {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        //add 3 for '.00' 
                        cellTextLength += (3 + thousandCount);
                    }

                    if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex)) {
                        //add an extra char for bold - not 100% acurate but good enough for what i need.
                        cellTextLength += 1;
                    }

                    if (maxColWidth.ContainsKey(i)) {
                        var current = maxColWidth[i];
                        if (cellTextLength > current) {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }

        Columns AutoSizeCells(SheetData sheetData) {
            var maxColWidth = GetMaxCharacterWidth(sheetData);

            Columns columns = new Columns();
            //this is the width of my font - yours may be different
            double maxWidth = 7;
            foreach (var item in maxColWidth) {
                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;
                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }

        ForegroundColor TranslateForeground(System.Drawing.Color fillColor) {
            return new ForegroundColor() {
                Rgb = new HexBinaryValue() {
                    Value =
                              System.Drawing.ColorTranslator.ToHtml(
                              System.Drawing.Color.FromArgb(
                                  fillColor.A,
                                  fillColor.R,
                                  fillColor.G,
                                  fillColor.B)).Replace("#", "")
                }
            };
        }

    }
}
