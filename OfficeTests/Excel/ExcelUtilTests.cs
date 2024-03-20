using Microsoft.VisualStudio.TestTools.UnitTesting;
using me.fengyj.CommonLib.Office.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace me.fengyj.CommonLib.Office.Excel.Tests {
    [TestClass()]
    public class ExcelUtilTests {

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

            using (SpreadsheetDocument excel = SpreadsheetDocument.Create("demo.xlsx", SpreadsheetDocumentType.Workbook)) {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = excel.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart1 = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart1.Worksheet = new Worksheet();
                SheetData data1 = CreateShee1Data();
                Columns columns1 = AutoSizeCells(data1);
                worksheetPart1.Worksheet.Append(columns1);
                worksheetPart1.Worksheet.Append(data1);
                DefineTable(worksheetPart1, 1, (uint)Sheet1Data.Length + 1, 1, (uint)Sheet1Headers.Length, Sheet1Headers);

                WorksheetPart worksheetPart2 = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart2.Worksheet = new Worksheet();
                SheetData data2 = CreateShee2Data();
                Columns columns2 = AutoSizeCells(data2);
                worksheetPart2.Worksheet.Append(columns2);
                worksheetPart2.Worksheet.Append(data2);

                var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateStyleSheet();
                stylesPart.Stylesheet.Save();

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
        void DefineTable(WorksheetPart worksheetPart, uint rowMin, uint rowMax, uint colMin, uint colMax, string[] headers) {

            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId" + (worksheetPart.TableDefinitionParts.Count() + 1));
            int tableNo = worksheetPart.TableDefinitionParts.Count();

            string reference = GetTableReference(rowMin, rowMax, colMin, colMax);

            Table table = new Table() { Id = (UInt32)tableNo, Name = "Table" + tableNo, DisplayName = "Table" + tableNo, Reference = reference, TotalsRowShown = false };
            AutoFilter autoFilter = new AutoFilter() { Reference = reference };

            TableColumns tableColumns = new TableColumns() { Count = (UInt32)(colMax - colMin + 1) };
            for (int i = 0; i < (colMax - colMin + 1); i++) {
                tableColumns.Append(new TableColumn() { Id = (UInt32)(colMin + i), Name = headers[i] }); //changed i+1 -> colMin + i
                                                                                                           //Add cell values (shared string)
            }

            TableStyleInfo tableStyleInfo = new TableStyleInfo() { Name = "TableStyleMedium1", ShowFirstColumn = true, ShowLastColumn = true, ShowRowStripes = true, ShowColumnStripes = false };

            table.Append(autoFilter);
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
            TablePart tablePart = new TablePart() { Id = "rId" + tableNo };

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
                StyleIndex = 8
            }, cellIndex);
        }

        void InsertTransitCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Number, CellValue = new CellValue(content), StyleIndex = 7 }, cellIndex);
        }

        void InsertNumberCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Number, CellValue = new CellValue(content), StyleIndex = 5 }, cellIndex);
        }

        void InsertDateCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Date, CellValue = new CellValue(content), StyleIndex = 1 }, cellIndex);
        }

        void InsertDateTimeCell(Row row, string content, int cellIndex) {
            row.InsertAt<Cell>(new Cell() { DataType = CellValues.Number, CellValue = new CellValue(DateTime.Parse(content).ToOADate().ToString()), StyleIndex = 3 }, cellIndex);
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

        Stylesheet CreateStyleSheet() {
            Stylesheet stylesheet = new Stylesheet();
            #region Number format
            uint DATETIME_FORMAT = 164;
            uint DIGITS4_FORMAT = 165;
            var numberingFormats = new NumberingFormats();
            numberingFormats.Append(new NumberingFormat // Datetime format
            {
                NumberFormatId = UInt32Value.FromUInt32(DATETIME_FORMAT),
                FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            });
            numberingFormats.Append(new NumberingFormat // four digits format
            {
                NumberFormatId = UInt32Value.FromUInt32(DIGITS4_FORMAT),
                FormatCode = StringValue.FromString("0000")
            });
            numberingFormats.Count = UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
            #endregion

            #region Fonts
            var fonts = new Fonts();
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font()  // Font index 0 - default
            {
                FontName = new FontName { Val = StringValue.FromString("Calibri") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) }
            });
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font()  // Font index 1
            {
                FontName = new FontName { Val = StringValue.FromString("Arial") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) },
                Bold = new Bold()
            });
            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
            #endregion

            #region Fills
            var fills = new Fills();
            fills.Append(new Fill() // Fill index 0
            {
                PatternFill = new PatternFill { PatternType = PatternValues.None }
            });
            fills.Append(new Fill() // Fill index 1
            {
                PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
            });
            fills.Append(new Fill() // Fill index 2
            {
                PatternFill = new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = TranslateForeground(System.Drawing.Color.LightBlue),
                    BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb }
                }
            });
            fills.Append(new Fill() // Fill index 3
            {
                PatternFill = new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = TranslateForeground(System.Drawing.Color.LightSkyBlue),
                    BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb }
                }
            });
            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            #endregion

            #region Borders
            var borders = new Borders();
            borders.Append(new Border   // Border index 0: no border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border    //Boarder Index 1: All
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border   // Boarder Index 2: Top and Bottom
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            #endregion

            #region Cell Style Format
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Append(new CellFormat  // Cell style format index 0: no format
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            });
            cellStyleFormats.Count = UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
            #endregion

            #region Cell format
            var cellFormats = new CellFormats();
            cellFormats.Append(new CellFormat());    // Cell format index 0
            cellFormats.Append(new CellFormat   // CellFormat index 1
            {
                NumberFormatId = 14,        // 14 = 'mm-dd-yy'. Standard Date format;
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 2: Standard Number format with 2 decimal placing
            {
                NumberFormatId = 4,        // 4 = '#,##0.00';
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell formt index 3
            {
                NumberFormatId = DATETIME_FORMAT,        // 164 = 'dd/mm/yyyy hh:mm:ss'. Standard Datetime format;
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 4
            {
                NumberFormatId = 3, // 3   #,##0
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 5
            {
                NumberFormatId = 4, // 4   #,##0.00
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 6
            {
                NumberFormatId = 10,    // 10  0.00 %,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 7
            {
                NumberFormatId = DIGITS4_FORMAT,    // Format cellas 4 digits. If less than 4 digits, prepend 0 in front
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 8: Cell header
            {
                NumberFormatId = 49,
                FontId = 1,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center }
            });
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            #endregion

            stylesheet.Append(numberingFormats);
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);

            #region Cell styles
            //var css = new CellStyles();
            //css.Append(new CellStyle {
            //    Name = StringValue.FromString("Normal"),
            //    FormatId = 0,
            //    BuiltinId = 0
            //});
            //css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            //stylesheet.Append(css);
            #endregion

            var dfs = new DifferentialFormats { Count = 0 };
            stylesheet.Append(dfs);
            var tss = new TableStyles {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            };
            stylesheet.Append(tss);

            return stylesheet;
        }

    }
}