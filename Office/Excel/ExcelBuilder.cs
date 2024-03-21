﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using me.fengyj.CommonLib.Utils;

namespace me.fengyj.CommonLib.Office.Excel {
    public class ExcelBuilder {

        public const int MAX_ROW_COUNT = 1_048_576;
        public List<SheetBuilder> SheetBuilders { get; private set; } = [];
        internal List<string> SheetNames { get; private set; } = [];
        internal List<string> TableNames { get; private set; } = [];
        internal List<string> TablesDefined { get; private set; } = [];
        internal uint RelationId { get; private set; } = 0;

        public SheetBuilder AppendSheet(string sheetName, bool autoColumnWidth = true, uint? maxColumnWidth = 120, Dictionary<int, uint>? columnWidths = null) {
            var sheetBuilder = new SheetBuilder(this, sheetName, autoColumnWidth: autoColumnWidth, maxColumnWidth: maxColumnWidth, columnWidths: columnWidths);
            this.SheetBuilders.Add(sheetBuilder);
            return sheetBuilder;
        }

        public void BuildTo(string filePath) {

            IOUtil.DeleteFile(filePath, true);

            using (var workbook = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook)) {

                var workbookPart = workbook.AddWorkbookPart();

                workbookPart.Workbook = new();
                workbookPart.Workbook.Sheets = new();

                this.SheetBuilders.ForEach(b => b.BuildTo(workbook));

                SheetStyleBuilder.BuildTo(workbookPart);

                workbook.Save();
                workbook.Dispose();
            }
        }

        public static void BuildTo(string filePath, DataSet dataset) {

            var builder = new ExcelBuilder();

            int sheetIdx = 1;
            foreach (DataTable tbl in dataset.Tables) {
                var sheetName = string.IsNullOrWhiteSpace(tbl.TableName) || tbl.TableName.Equals("Table", StringComparison.OrdinalIgnoreCase)
                    ? $"Sheet{sheetIdx++}"
                    : tbl.TableName;
                builder.AppendSheet(sheetName).AddTable(tbl);
            }

            builder.BuildTo(filePath);
        }

        internal string TryAddOrGetSheetName(string? name) {

            if (string.IsNullOrWhiteSpace(name)) name = $"Sheet{this.SheetNames.Count + 1}";
            return StringUtil.TryAddOrGetNewNameIfDuplicated(this.SheetNames, name, maxLength: 31); // the sheet name cannot longer than 31.
        }

        internal string TryAddOrGetTableName(string? name) {

            if (string.IsNullOrWhiteSpace(name)) name = $"Table{this.TableNames.Count + 1}";
            name = name.Replace(' ', '_');
            return StringUtil.TryAddOrGetNewNameIfDuplicated(this.TableNames, name);
        }

        internal uint DefineTableAndGetId(string tableName) {
            this.TablesDefined.Add(tableName);
            return (uint)this.TablesDefined.Count;
        }

        internal string GetNewRelationId() {
            return $"rId{++RelationId}";
        }

        public static string GetExcelColumnName(string prefix, uint column) {
            return column < 26
                ? $"{prefix}{(char)('A' + column - 1)}"
                : GetExcelColumnName(GetExcelColumnName(prefix, (column - column % 26) / 26 - 1), column % 26);
        }

        public static string GetExcelCellReference(uint row, uint column) => $"{GetExcelColumnName(string.Empty, column)}{row}";

        public static string GetExcelTableReference(uint rowStart, uint rowEnd, uint columnStart, uint columnEnd)
            => $"{GetExcelCellReference(rowStart, columnStart)}:{GetExcelCellReference(rowEnd, columnEnd)}";
    }

    public class SheetBuilder {

        private static readonly Func<IEnumerable<RowBuilder>> EmptyRowsBuilder = () => [];

        public SheetBuilder(ExcelBuilder excelBuilder, string? sheetName, bool autoColumnWidth = true, uint? maxColumnWidth = 120, Dictionary<int, uint>? columnWidths = null) {

            this.ExcelBuilder = excelBuilder;
            this.SheetNameForBuild = excelBuilder.TryAddOrGetSheetName(sheetName);
            this.SheetName = sheetName ?? this.SheetNameForBuild;
            this.AutoColumnWidth = autoColumnWidth;
            this.MaxColumnWidth = maxColumnWidth;
            this.ColumnWidths = columnWidths;

            this.HyperlinkBuilder = new HyperlinkBuilder(this);

            this.RowBuilderSuppliers = EmptyRowsBuilder;
        }

        public ExcelBuilder ExcelBuilder { get; private set; }
        public string SheetName { get; private set; }
        public string SheetNameForBuild { get; private set; }
        public bool AutoColumnWidth { get; private set; }
        public uint? MaxColumnWidth { get; private set; }
        public Dictionary<int, uint>? ColumnWidths { get; private set; }
        public uint CurrentRowPosition { get; internal set; }
        public Func<IEnumerable<RowBuilder>> RowBuilderSuppliers { get; private set; }
        public List<TableBuilder> TableBuilders { get; private set; } = [];
        public HyperlinkBuilder HyperlinkBuilder { get; private set; }

        public SheetBuilder AddRow(Func<RowBuilder> rowBuilder) {
            return AddRows(() => Enumerable.Repeat(rowBuilder(), 1));
        }

        public SheetBuilder AddRow(IEnumerable<object> values, CellStyle? style = null, uint rowOffset = 1, uint colOffset = 1) {
            return AddRow(() => CreateRowBuilder(rowOffset: rowOffset).AddCells(values, cellStyle: style, colOffset: colOffset));
        }

        public SheetBuilder AddRow(object value, CellStyle? style = null, uint rowOffset = 1, uint colOffset = 1) {
            return AddRow(() => CreateRowBuilder(rowOffset: rowOffset).AddCell(value, cellStyle: style, colOffset: colOffset));
        }

        public SheetBuilder AddRows(Func<IEnumerable<RowBuilder>> rowsBuilder) {

            if (this.RowBuilderSuppliers == EmptyRowsBuilder) {
                this.RowBuilderSuppliers = rowsBuilder;
            }
            else {
                var orig = this.RowBuilderSuppliers;
                this.RowBuilderSuppliers = () => Enumerable.Concat(orig(), rowsBuilder());
            }
            return this;
        }

        public SheetBuilder AddRows(IEnumerable<object> values, uint rowOffset = 1, uint colOffset = 1) {

            return AddRows(() => values.Select((r, idx) => {
                var rb = CreateRowBuilder(rowOffset: idx == 0 ? rowOffset : 1);
                if (r is IEnumerable<object> cs) return rb.AddCell(cs, colOffset: colOffset);
                else return rb.AddCell(r, colOffset: colOffset);
            }));
        }

        public SheetBuilder AddTable<T>(IEnumerable<T> data, TableConfig<T> config, uint rowOffset = 1, uint colOffset = 1) {

            if (config == null)
                throw new ArgumentNullException(nameof(config));

            config.Columns.ForEach(c => {
                if (!c.HasDataGetter)
                    throw new ArgumentException($"The column {c.ColumnName} doesn't set DataGetter.", nameof(config));
            });

            if (data == null)
                throw new ArgumentNullException(nameof(data));

            AddRows(() => CreateTable(
                data.Select(i => config.Columns.Select(c => c.GetDataObject(i))),
                config: config,
                rowOffset: rowOffset,
                colOffset: colOffset));

            return this;
        }

        public SheetBuilder AddTable(
            DataTable dataTable,
            TableStyle? tableStyle = null,
            Dictionary<string, CellStyle>? cellStyles = null,
            HashSet<string>? colsInclude = null,
            HashSet<string>? colsExclude = null,
            uint rowOffset = 1,
            uint colOffset = 1) {

            var columns = new List<ITableColumnConfig<DataRow>>();
            int colIdx = -1;
            foreach (DataColumn c in dataTable.Columns) {

                colIdx++;
                if ((colsInclude?.Contains(c.ColumnName) ?? true) && !(colsExclude?.Contains(c.ColumnName) ?? false)) {

                    var idx = colIdx;
                    columns.Add(new TableColumnConfig<DataRow, object>(
                        c.ColumnName,
                        style: (cellStyles?.ContainsKey(c.ColumnName) ?? false) ? cellStyles[c.ColumnName] : null,
                        dataType: c.DataType,
                        dataGetter: i => i[idx] == DBNull.Value ? null : i[idx]));
                }
            }

            var tblConfig = new TableConfig<DataRow>(columns, style: tableStyle);
            var data = GetTableData(dataTable);
            return AddTable(data, config: tblConfig, rowOffset: rowOffset, colOffset: colOffset);
        }

        public void BuildTo(SpreadsheetDocument workbook) {

            if (workbook.WorkbookPart == null)
                throw new ArgumentException("The WorkbookPart hasn't been initialized.", nameof(workbook));

            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            sheetPart.Worksheet = new Worksheet();
            var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
                throw new ArgumentException("The Sheets hasn't been initialized.", nameof(workbook));

            var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
            var sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId?.Value).DefaultIfEmpty(0u).Max() + 1;
            var sheetName = this.SheetNameForBuild;

            var sheet = new Sheet { Id = relationshipId, SheetId = sheetId, Name = SheetName };
            sheets.Append(sheet);

            var sheetData = new SheetData();
            
            this.AddRow(string.Empty); // add a empty cell to the end to avoid the error caused by the empty sheet.
            if (this.RowBuilderSuppliers != null) {

                var builders = this.RowBuilderSuppliers();
                if (builders != null) {
                    var missedRows = 0;
                    RowBuilder? lastBuilder = null;
                    foreach (var b in builders) {
                        lastBuilder = b;
                        if (this.CurrentRowPosition >= ExcelBuilder.MAX_ROW_COUNT) missedRows++;
                        else b.BuildTo(sheetData);
                    }
                    if (missedRows > 1)
                        new RowBuilder(this).AddCell($"(There are {missedRows} more rows cannot be displayed during to the limitaion of Excel.)").BuildTo(sheetData);
                    else
                        lastBuilder?.BuildTo(sheetData);
                }

                SetColumnWidths(sheetPart, sheetData);
            }

            // todo: DO NOT change the order of code below.

            sheetPart.Worksheet.Append(sheetData);

            this.HyperlinkBuilder.BuildTo(sheetPart);

            this.TableBuilders.ForEach(tb => tb.BuildTo(sheetPart));
        }

        private IEnumerable<RowBuilder> CreateTable<T>(IEnumerable<IEnumerable<object?>> data, TableConfig<T>? config = null, uint rowOffset = 1, uint colOffset = 1) {

            var startRow = this.CurrentRowPosition + rowOffset;
            var startCol = colOffset;
            var endCol = startCol;

            // 1. table header
            var hasHeader = config != null && config.TableStyle.ShowHeader && config.Columns != null && config.Columns.Any(c => c.ColumnName != null);
            if (hasHeader && config != null && config.Columns != null) { // check config and columns again to avoid the warning in below code

                endCol = (uint)(startCol - 1 + config.Columns.Count);

                yield return CreateRowBuilder(rowOffset: rowOffset).AddCells(
                    config.Columns.Select((c, idx) => c.ColumnName ?? $"Column {idx + 1}"),
                    cellStyle: CellStyle.TableHeader,
                    colOffset: colOffset);
            }

            // 2. data
            RowBuilder? lastRowBuilder = null;
            if (data != null) {

                var ro = hasHeader ? 1 : rowOffset;
                foreach (var r in data) {
                    lastRowBuilder = CreateRowBuilder(rowOffset: ro);
                    lastRowBuilder.AddCells(() => r.Select((i, idx) => new CellBuilder(lastRowBuilder, i, config?.Columns?[idx].CellStyle) {
                        ColumnOffset = idx == 0 ? colOffset : 1
                    }));
                    yield return lastRowBuilder;
                    ro = 1; // for rest rows the offset should be 1.
                }
                if (!hasHeader && lastRowBuilder != null)
                    endCol = Math.Max(endCol, lastRowBuilder.CurrentColumnPosition);
            }

            if (hasHeader && lastRowBuilder == null) { // means no data, need to create an empty row to avoid the error to create the table

                lastRowBuilder = CreateRowBuilder(rowOffset: 1);
                lastRowBuilder.AddCells(new object[endCol - startCol + 1], colOffset: colOffset);
                yield return lastRowBuilder;
            }

            if (lastRowBuilder != null) { // if no data nor header, don't create the table.
                this.TableBuilders.Add(new TableBuilder(
                    this,
                    startRow,
                    this.CurrentRowPosition,
                    startCol,
                    endCol,
                    tableName: config?.TableName,
                    headers: hasHeader ? config?.Columns?.Select(i => i.ColumnName).ToArray() : null,
                    style: config?.TableStyle));
            }
        }

        private IEnumerable<DataRow> GetTableData(DataTable dataTable) {
            foreach (DataRow row in dataTable.Rows) {
                yield return row;
            }
        }

        private RowBuilder CreateRowBuilder(uint rowOffset = 1) {
            return new RowBuilder(this) { RowOffset = rowOffset };
        }

        private void SetColumnWidths(WorksheetPart workSheetPart, SheetData sheetData) {

            var columns = new Columns();
            var maxLengthOfEachCol = GetMaxLengthForEachColumn(sheetData);

            for (var i = 1; i <= maxLengthOfEachCol.Count; i++) {
                var w = 0u;
                if (this.ColumnWidths != null && this.ColumnWidths.ContainsKey(i))
                    w = this.ColumnWidths[i];
                else
                    w = maxLengthOfEachCol[i - 1];
                w = Math.Min(Math.Max(0, w), this.MaxColumnWidth ?? 0);
                var width = Math.Truncate((w * 7.0 + 5) / 7.0 * 256) / 256;
                var col = new Column { BestFit = this.AutoColumnWidth, Min = (uint)i, Max = (uint)i, CustomWidth = true, Width = DoubleValue.FromDouble(width) };

                columns.Append(col);
            }

            workSheetPart.Worksheet.Append(columns);
        }

        private static List<uint> GetMaxLengthForEachColumn(SheetData sheetData) {

            var lengths = new Dictionary<int, uint>();
            var maxRowsToCheck = 100; // avoid to cost too much time on it
            int rowIdx = 0;

            foreach (var row in sheetData.Elements<Row>()) {

                if (rowIdx++ > maxRowsToCheck) break;

                var cells = row.Elements<Cell>().ToArray();
                for (var colIdx = 0; colIdx < cells.Length; colIdx++) {
                    var cell = cells[colIdx];
                    var val = cell.CellValue?.InnerText ?? cell.InnerText;
                    var length = (uint)(val?.Length ?? 2);
                    if (cell.DataType?.Value == CellValues.Date || cell.DataType?.Value == CellValues.Number)
                        length = (uint)(Math.Ceiling(length * 1.3));
                    if (lengths.TryGetValue(colIdx, out var value)) {
                        var l = value;
                        lengths[colIdx] = Math.Max(l, length);
                    }
                    else {
                        lengths[colIdx] = length;
                    }
                }
            }

            return lengths.OrderBy(i => i.Key)
                .Select(i => i.Value)
                .Select(i => {
                    if (i < 5) return 12;
                    else if (i < 10) return i * 1.5;
                    else if (i < 25) return i * 1.2;
                    else if (i < 40) return i * 1.1;
                    else return i * 1.05;
                })
                .Select(i => (uint)(Math.Truncate(i)))
                .ToList();
        }
    }

    public class RowBuilder {

        private Func<IEnumerable<CellBuilder>> EmptyCellsBuilder = () => [];
        public RowBuilder(SheetBuilder sheetBuilder) {
            this.SheetBuilder = sheetBuilder;
            this.CellBuilderSuppliers = EmptyCellsBuilder;
        }

        public SheetBuilder SheetBuilder { get; private set; }
        public uint RowOffset { get; set; }
        public uint CurrentColumnPosition { get; internal set; } = 0;
        public Func<IEnumerable<CellBuilder>> CellBuilderSuppliers { get; private set; }

        public RowBuilder AddCells(Func<IEnumerable<CellBuilder>> cellBuilderSupplier) {
            if (this.CellBuilderSuppliers == EmptyCellsBuilder) {
                this.CellBuilderSuppliers = cellBuilderSupplier;
            }
            else {
                var orig = this.CellBuilderSuppliers;
                this.CellBuilderSuppliers = () => Enumerable.Concat(orig(), cellBuilderSupplier());
            }
            return this;
        }

        public RowBuilder AddCells(IEnumerable<object> values, CellStyle? cellStyle = null, uint colOffset = 1) {

            if (values == null) return this;

            return AddCells(() => values.Select((v, idx) => {
                var cb = new CellBuilder(this, v, style: cellStyle);
                if (idx == 0) cb.ColumnOffset = colOffset;
                return cb;
            }));
        }

        public RowBuilder AddCell(CellBuilder cellBuilder) {
            return AddCells(() => Enumerable.Repeat(cellBuilder, 1));
        }

        public RowBuilder AddCell(object cellValue, CellStyle? cellStyle = null, uint colOffset = 1) {
            return AddCells(() => Enumerable.Repeat(cellValue, 1).Select(v => {
                var cb = new CellBuilder(this, v, style: cellStyle);
                cb.ColumnOffset = colOffset;
                return cb;
            }));
        }

        public void BuildTo(SheetData sheetData) {

            for (var i = 1; i < this.RowOffset; i++)
                sheetData.AppendChild(new Row());

            var row = new Row();

            this.SheetBuilder.CurrentRowPosition += this.RowOffset;
            sheetData.AppendChild(row);

            var builders = this.CellBuilderSuppliers();
            if (builders != null)
                foreach (var builder in builders)
                    builder.BuildTo(row);
        }
    }

    public class CellBuilder {

        private static readonly Func<object, Cell> defaultCellGetter = obj => new Cell { InlineString = new() { Text = new Text(obj.ToString() ?? string.Empty) } };
        private static readonly Dictionary<Type, Tuple<CellStyle, Func<object, Cell>>> DefaultCellBuilders = new() {

            { typeof(string), System.Tuple.Create(CellStyle.Cell, defaultCellGetter)},

            { typeof(short), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((short)obj)}) },
            { typeof(ushort), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((ushort)obj)}) },
            { typeof(int), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((int)obj)}) },
            { typeof(uint), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((int)(uint)obj)}) },
            { typeof(byte), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((byte)obj)}) },
            { typeof(sbyte), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((sbyte)obj)}) },

            { typeof(long), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((double)(long)obj)}) },
            { typeof(ulong), System.Tuple.Create(CellStyle.Cell_Integer_Default, (object obj) => new Cell{CellValue = new ((double)(ulong)obj)}) },

            { typeof(float), System.Tuple.Create(CellStyle.Cell_Decimal_Default, (object obj) => new Cell{CellValue = new ((double)(float)obj)}) },
            { typeof(double), System.Tuple.Create(CellStyle.Cell_Decimal_Default, (object obj) => new Cell{CellValue = new ((double)obj)}) },
            { typeof(decimal), System.Tuple.Create(CellStyle.Cell_Decimal_Default, (object obj) => new Cell{CellValue = new ((decimal)obj)}) },

            { typeof(bool), System.Tuple.Create(CellStyle.Cell_Bool_Default, (object obj) => new Cell{CellValue = new ((bool)obj)}) },

            { typeof(DateTime), System.Tuple.Create(CellStyle.Cell_DateTime_Default, (object obj) => new Cell{CellValue = new ((DateTime)obj)}) },

            { typeof(TimeSpan), System.Tuple.Create(CellStyle.Cell_Timespan_Default, (object obj) => new Cell{InlineString = new() { Text = new Text(((TimeSpan)obj).ToString() ?? string.Empty) }}) },

            { typeof(HyperLinkValue), System.Tuple.Create(CellStyle.Hyperlink, (object obj) => new Cell{InlineString = new() { Text = new Text(((HyperLinkValue)obj)?.DisplayName ?? string.Empty) }  }) },
        };

        public CellBuilder(RowBuilder rowBuilder, object? cellValue, CellStyle? style = null) {

            this.RowBuilder = rowBuilder;
            this.CellValue = cellValue;
            this.Cell = BuildCell(cellValue, style);
        }

        public uint ColumnOffset { get; set; } = 1;
        public RowBuilder RowBuilder { get; private set; }
        public Cell Cell { get; private set; }
        private object? CellValue { get; set; }

        public void BuildTo(Row row) {

            for (var i = 1; i < this.ColumnOffset; i++)
                row.AppendChild(new Cell());

            this.RowBuilder.CurrentColumnPosition += this.ColumnOffset;

            if(this.CellValue is HyperLinkValue linkValue) {

                if (!string.IsNullOrWhiteSpace(linkValue.Link)) {

                    linkValue.CellReferenceId = ExcelBuilder.GetExcelCellReference(
                        this.RowBuilder.SheetBuilder.CurrentRowPosition,
                        this.RowBuilder.CurrentColumnPosition);

                    this.RowBuilder.SheetBuilder.HyperlinkBuilder.HyperlinkValues.Add(linkValue);
                }
            }

            row.AppendChild(this.Cell);
        }

        private Cell BuildCell(object? obj, CellStyle? style) {

            var cellStyle = style ?? CellStyle.Cell;

            if (obj == null) return new Cell() { StyleIndex = cellStyle.StyleId, DataType = cellStyle.CellValueType };

            object val = obj;
            var cellGetter = defaultCellGetter;

            if (DefaultCellBuilders.TryGetValue(val.GetType(), out var item)) {
                cellStyle = style ?? item.Item1;
                cellGetter = item.Item2;
            }

            var cell = cellGetter(val);

            cell.StyleIndex = cellStyle.StyleId;
            cell.DataType = cellStyle.CellValueType;

            return cell;
        }
    }

    public class TableBuilder {

        public TableBuilder(
            SheetBuilder sheetBuilder,
            uint rowStart,
            uint rowEnd,
            uint colStart,
            uint colEnd,
            string? tableName = null,
            string[]? headers = null,
            TableStyle? style = null) {

            this.SheetBuilder = sheetBuilder;
            this.RowStart = rowStart;
            this.RowEnd = rowEnd;
            this.ColumnStart = colStart;
            this.ColumnEnd = colEnd;
            this.Style = style ?? new TableStyle();

            this.HasTotalRow = false; // not supported for now.

            if (headers != null && headers.Length != colEnd - colStart + 1)
                throw new ArgumentException("The headers' length doesn't match the table's width.");

            this.TableName = this.SheetBuilder.ExcelBuilder.TryAddOrGetTableName(tableName);

            this.Headers = headers ?? new string[colEnd - colStart + 1];
            this.HasHeader = style?.ShowHeader ?? headers != null;
            for (int i = 0; i < this.Headers.Length; i++)
                if (string.IsNullOrWhiteSpace(this.Headers[i])) this.Headers[i] = $"Column {i + 1}";
        }

        public uint Id { get; set; }
        public SheetBuilder SheetBuilder { get; private set; }
        public string TableName { get; private set; }
        public uint RowStart { get; private set; }
        public uint RowEnd { get; private set; }
        public uint ColumnStart { get; private set; }
        public uint ColumnEnd { get; private set; }
        public string[] Headers { get; private set; }
        public bool HasTotalRow { get; private set; }
        public bool HasHeader { get; private set; }
        public TableStyle Style { get; private set; }

        public void BuildTo(WorksheetPart worksheetPart) {

            // DO NOT CHANGE the order of adding following objects to the worksheetPart. it will cause some errors to the generated file.

            this.Id = this.SheetBuilder.ExcelBuilder.DefineTableAndGetId(this.TableName);

            var tblRefId = this.SheetBuilder.ExcelBuilder.GetNewRelationId();
            var tblDefPart = worksheetPart.AddNewPart<TableDefinitionPart>(tblRefId);
            var tblRef = ExcelBuilder.GetExcelTableReference(this.RowStart, this.RowEnd, this.ColumnStart, this.ColumnEnd);
            var tbl = new Table { Id = this.Id, Name = this.TableName, DisplayName = this.TableName, Reference = tblRef, TotalsRowShown = HasTotalRow };
            
            var cols = new TableColumns { Count = UInt32Value.FromUInt32((uint)this.Headers.Length) };
            for (int i = 0; i < this.Headers.Length; i++) cols.Append(new TableColumn { Id = (uint)i + this.ColumnStart, Name = this.Headers[i].Trim() });

            if (HasHeader)
                tbl.Append(new AutoFilter { Reference = tblRef });

            tbl.Append(cols);

            if (!this.HasHeader) 
                tbl.HeaderRowCount = UInt32Value.FromUInt32(0);

            tbl.Append(this.Style.TableStyleInfo);

            tblDefPart.Table = tbl;

            var tblParts = worksheetPart.Worksheet.ChildElements.Where(e => e is TableParts).FirstOrDefault();
            if (tblParts == null) {
                tblParts = new TableParts();
                worksheetPart.Worksheet.Append(tblParts);
            }

            var tblPart = new TablePart { Id = tblRefId };
            tblParts.Append(tblPart);
        }
    }

    public class HyperlinkBuilder {

        public HyperlinkBuilder(SheetBuilder sheetBuilder) {
            this.SheetBuilder = sheetBuilder;
        }

        public SheetBuilder SheetBuilder { get; private set; }
        public List<HyperLinkValue> HyperlinkValues { get; set; } = [];

        public void BuildTo(WorksheetPart worksheetPart) {

            if (this.HyperlinkValues == null || this.HyperlinkValues.Count == 0) return;

            var links = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault() ?? new Hyperlinks();

            this.HyperlinkValues.ForEach(v => {

                var rId = this.SheetBuilder.ExcelBuilder.GetNewRelationId();

                var link = new Hyperlink { Id = StringValue.FromString(rId), Reference = StringValue.FromString(v.CellReferenceId) };
                worksheetPart.AddHyperlinkRelationship(new Uri(v.Link), true, rId);

                links.AppendChild(link);
            });

            worksheetPart.Worksheet.Append(links);
        }
    }
}
