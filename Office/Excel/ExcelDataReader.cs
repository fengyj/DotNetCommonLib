﻿using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace me.fengyj.CommonLib.Office.Excel {
    public class ExcelDataReader<T> : IDisposable {

        private static readonly Cell EmptyCell = new();

        private SpreadsheetDocument doc;
        private SharedStringItem[]? sharedStrings;
        private int isDisposed = 0;

        public ExcelDataReader(string filePath) {
            // if OpenXmlPackageException was thrown, there is a bug that the file is opened and not released. 
            // then the file cannot be changed till the process is closed.
            this.doc = SpreadsheetDocument.Open(filePath, false);
        }

        public IEnumerable<Record> Read(
            Config config,
            T? reused = default) {

            var sheet = this.doc.WorkbookPart?.Workbook.Sheets?.Skip((int)config.SheetNo - 1).FirstOrDefault();
            if (sheet is not Sheet s) yield break;
            var sheetPart = this.doc.WorkbookPart?.GetPartById(s.Id?.Value ?? "");
            if (sheetPart == null) yield break;

#pragma warning disable CS8629 // Nullable value type may be null.
            var dict = config.Deserializers.Where(i => i.ColumnIndex.HasValue).GroupBy(i => i.ColumnIndex.Value).ToDictionary(i => i.Key, i => i.ToList());
#pragma warning restore CS8629 // Nullable value type may be null.

            var rec = new Record();

            using var reader = OpenXmlReader.Create(sheetPart);
            var currentRowIdx = 0u;
            var currentColIdx = 0u;
            var allCellsInRowEmpty = true;

            while (reader.Read()) {

                if (reader.ElementType == typeof(Row)) {

                    if (reader.IsStartElement) {

                        allCellsInRowEmpty = true;
                        var rowId = GetRowId(reader);
                        if (rowId != null)
                            for (var i = currentRowIdx + 1; i < rowId.Value; i++)
                                currentRowIdx++;

                        currentRowIdx++;
                        currentColIdx = 0; // reset col idx, because it's a new row now.
                        var rowCanReturn = config.IsFirstLineHeader ? (currentRowIdx > config.Area.IndexOfRowBegin) : (currentRowIdx >= config.Area.IndexOfRowBegin);
                        if (rowCanReturn)
                            rec = (rec ?? new Record()).GetReusedOrCreateNew(currentRowIdx, reused, () => config.RecordBuilder(currentRowIdx));
                    }
                    else {
                        // check if needs to skip the header row
                        var rowCanReturn = config.IsFirstLineHeader ? (currentRowIdx > config.Area.IndexOfRowBegin) : (currentRowIdx >= config.Area.IndexOfRowBegin);
                        if (rowCanReturn && currentColIdx >= config.Area.IndexOfColumnBegin)
                            if (rec != null && (!config.SkipEmptyRows || !allCellsInRowEmpty))
                                yield return rec;

                        if (config.Area.IndexOfRowEnd.HasValue && currentRowIdx >= config.Area.IndexOfRowEnd.Value) break;
                    }
                }
                else if (reader.ElementType == typeof(Cell) && reader.IsStartElement && rec != null) {

                    var el = reader.LoadCurrentElement();
                    if (el is not Cell) continue;
                    var cell = el as Cell;

#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    var cellRef = cell.CellReference?.Value;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    var colId = GetCellColumn(cellRef);
                    if (colId.HasValue) {
                        // in case the cells are not filled to each column
                        while (currentColIdx < colId.Value - 1) {

                            currentColIdx++;

                            if (currentRowIdx < config.Area.IndexOfRowBegin || currentColIdx < config.Area.IndexOfColumnBegin) continue;
                            if (config.Area.IndexOfColumnEnd.HasValue && currentColIdx > config.Area.IndexOfColumnEnd.Value) continue;

                            if (dict.TryGetValue(currentColIdx, out var deserializers)) {
                                deserializers.ForEach(d => d.DeserializeTo(rec, EmptyCell, this));
                            }
                            else if (config.DefaultSetter != null && rec.Data != null) {
                                config.DefaultSetter(rec.Data, currentColIdx, this.GetStringCellValue(EmptyCell));
                            }
                        }
                    }

                    currentColIdx++;

                    if (currentRowIdx < config.Area.IndexOfRowBegin || currentColIdx < config.Area.IndexOfColumnBegin) continue;
                    if (config.Area.IndexOfColumnEnd.HasValue && currentColIdx > config.Area.IndexOfColumnEnd.Value) continue;

                    if (config.IsFirstLineHeader && currentRowIdx == config.Area.IndexOfRowBegin) {
                        var name = cell.CellValue?.InnerText;
                        var matched = config.Deserializers.Where(d => d.ColumnName == name && !d.ColumnIndex.HasValue);
                        foreach (var d in matched) {
                            d.ColumnIndex = currentColIdx;
                            d.ColumnName = ExcelUtil.GetColumnName(currentColIdx);
                            if (dict.TryGetValue(currentColIdx, out var lst)) lst.Add(d);
                            else dict.Add(currentColIdx, [d]);
                        }
                    }
                    else if (dict.TryGetValue(currentColIdx, out var deserializers)) {
                        deserializers.ForEach(d => d.DeserializeTo(rec, cell, this));
                    }
                    else if (config.DefaultSetter != null && rec.Data != null) {
                        config.DefaultSetter(rec.Data, currentColIdx, this.GetStringCellValue(cell));
                    }

                    if (allCellsInRowEmpty && !string.IsNullOrWhiteSpace(this.GetCellValue(cell)))
                        allCellsInRowEmpty = false;
                }
            }

            config.Area.IndexOfRowEnd = currentRowIdx;
        }

        public string? GetCellValue(Cell cell) {

            if (cell.DataType?.Value == CellValues.SharedString) {
                return int.TryParse(cell.CellValue?.InnerText, out var sid) ? this.GetSharedString(sid) : null;
            }
            else if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText;
            }
            else if (cell == EmptyCell || cell.CellValue == null || cell.CellValue.InnerText == null) {
                return null;
            }
            else {
                return cell.CellValue.InnerText;
            }
        }

        public string? GetStringCellValue(Cell cell) {

            if (cell.DataType?.Value == CellValues.SharedString) {
                return int.TryParse(cell.CellValue?.InnerText, out var sid) ? this.GetSharedString(sid) : null;
            }
            else if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText;
            }
            else if (cell == EmptyCell || cell.CellValue == null || cell.CellValue.InnerText == null) {
                return null;
            }
            else if (double.TryParse(cell.CellValue.InnerText, out var v) && double.IsFinite(v)) {
                return Convert.ToDecimal(v).ToString();
            }
            else {
                return cell.CellValue.InnerText;
            }
        }

        private string? GetSharedString(int sid) {

            if (this.sharedStrings == null) {

                var shareTablePart = this.doc.WorkbookPart?.SharedStringTablePart;
                if (shareTablePart != null) {
                    this.sharedStrings = shareTablePart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                }
                if (this.sharedStrings == null)
                    this.sharedStrings = [];
            }
            return sid >= this.sharedStrings.Length ? null : this.sharedStrings[sid].InnerText;
        }

        private static uint GetSheetColumnsCount(Columns cols) {

            var items = cols.Elements<Column>()?.ToList();
            if (items == null) return 0;
            var maxColIdx = cols.Elements<Column>()?.Select(c => c.Max?.Value).Where(c => c != null).Max();
            return Math.Max((uint)items.Count, maxColIdx ?? (uint)items.Count);
        }

        private static uint? GetRowId(OpenXmlReader reader) {

            if (!reader.HasAttributes) return default(uint?);
            return reader.Attributes.Where(a => a.LocalName == "r")
                .Select(a => uint.TryParse(a.Value, out var id) ? id : default(uint?))
                .FirstOrDefault(default(uint?));
        }

        private static uint? GetCellColumn(string? cellRef) {
            if (cellRef == null) return default(uint?);
            var (_, c) = ExcelUtil.GetCellRowAndColumnIndex(cellRef);
            return c;
        }

        public void Dispose() {

            if (Interlocked.Exchange(ref this.isDisposed, 1) == 0) {
                try {
                    this.doc.Dispose();
                }
                catch { }
            }
        }

        /// <summary>
        /// Definitions about how to read the data from a specificed sheet.
        /// </summary>
        public class Config {

            /// <summary>
            /// 
            /// </summary>
            /// <param name="sheetNo">Which sheet contains the data to read. Starts from 1.</param>
            /// <param name="area"> Define the range of the cells to read. the end or row and/or the end of column can be omitted.</param>
            /// <param name="recBuilder">The builder function to create an object resprents the row in the sheet. The arg is the current row index.</param>
            /// <param name="deserializers">Define the functions how to convert cell value to the values in the object.</param>
            /// <param name="defaultSetter">If necessary, when deserializers don't contain the function to handle the cell value for a particular column, can set a default one.
            /// The second arg is the current col index.
            /// </param>
            /// <param name="skipEmptyRows">Skip the empty rows (all the cells are empty) or not. If yes, the row won't yield a T object.</param>
            /// <param name="isFirstLineHeader">
            /// Indicate if needs to skip the first record is table header.
            /// If any deserializer only specifics the table column name (the parameter tblColName), the first line in the area has to be table header.
            /// </param>
            public Config(
                uint sheetNo,
                DataArea area,
                Func<uint, T> recBuilder,
                List<DataDeserializer>? deserializers = null,
                Action<T, uint, string?>? defaultSetter = null,
                bool skipEmptyRows = true,
                bool isFirstLineHeader = false) {

                this.SheetNo = sheetNo;
                this.Area = area;
                this.Deserializers = deserializers ?? [];
                this.RecordBuilder = recBuilder;
                this.DefaultSetter = defaultSetter;
                this.SkipEmptyRows = skipEmptyRows;
                this.IsFirstLineHeader = isFirstLineHeader
                    || (deserializers?.Any(i => i.TableColumnName != null && i.ColumnIndex == null && i.ColumnName == null) ?? false);
            }

            /// <summary>
            /// The index of the sheet contains the data to read. Starts from 1.
            /// </summary>
            public uint SheetNo { get; private set; }
            public DataArea Area { get; private set; }
            public List<DataDeserializer> Deserializers { get; private set; }
            public Func<uint, T> RecordBuilder { get; private set; }
            public Action<T, uint, string?>? DefaultSetter { get; private set; }
            public bool SkipEmptyRows { get; private set; }
            public bool IsFirstLineHeader { get; private set; }
        }

        public abstract class DataDeserializer {

            private uint? colIdx;
            private string? colName;

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DataDeserializer(uint? colIdx = null, string? colName = null, string? tblColName = null, Func<string, bool>? equalsToNullChecker = null) {

                if (colIdx == null && colName == null && tblColName == null)
                    throw new ArgumentException("At least one of the arguments should have value.");

                if (colIdx.HasValue) this.ColumnIndex = colIdx.Value;
                if (colName != null) this.ColumnName = colName;
                this.TableColumnName = tblColName;
                this.EqualsToNullChecker = equalsToNullChecker;
            }

            public uint? ColumnIndex {
                get { return this.colIdx; }
                set {
                    this.colIdx = value;
                    this.colName = value.HasValue ? ExcelUtil.GetColumnName(value.Value) : null;
                }
            }

            public string? ColumnName {
                get { return this.colName; }
                set {
                    this.colName = value;
                    this.colIdx = value != null ? ExcelUtil.GetColumnIndex(value) : null;
                }
            }

            public string? TableColumnName { get; set; }
            public Func<string, bool>? EqualsToNullChecker { get; set; }

            public abstract void DeserializeTo(Record rec, Cell cell, ExcelDataReader<T> reader);

            protected Action<Record, V> GetValueSetterWrapper<V>(Action<T, V> action) {
                return (rec, v) => {
                    try {
                        if (rec.Data == null)
                            rec.AddCellError(this.colName, "Data hasn't been initialized.");
                        else
                            action(rec.Data, v);
                    }
                    catch (Exception ex) {
                        rec.AddCellError(this.colName, ex.Message);
                    }
                };
            }

            protected void SetErrorMsgAndWithDefaultValue<V>(Record rec, string errMsg, V defaultValue, Action<Record, V> setter) {
                setter(rec, defaultValue);
                if (defaultValue == null)
                    rec.AddCellError(this.colName, errMsg);
                else
                    rec.AddCellError(this.colName, $"[Warning] {errMsg} Set with default value {defaultValue}.");
            }

            protected string GetStringCellValue(Cell cell, ExcelDataReader<T> reader) {
                var str = reader.GetStringCellValue(cell);
                if (str == null || this.EqualsToNullChecker != null && this.EqualsToNullChecker(str)) return null;
                else return str;
            }

            protected string GetCellValue(Cell cell, ExcelDataReader<T> reader) {
                var str = reader.GetCellValue(cell);
                if (str == null || this.EqualsToNullChecker != null && this.EqualsToNullChecker(str)) return null;
                else return str;
            }
        }

        public class TextDeserializer : DataDeserializer {

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public TextDeserializer(
                Action<T, string?> valueSetter,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                string? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.DefaultValue = defaultValue;
            }

            public Action<Record, string?> ValueSetter { get; private set; }
            public string? DefaultValue { get; private set; }
            public override void DeserializeTo(Record rec, Cell cell, ExcelDataReader<T> reader) {

                var val = this.GetStringCellValue(cell, reader);
                if (rec.Data != null)
                    this.ValueSetter(rec, val ?? this.DefaultValue);
                else
                    rec.Error = "Data is not initialized.";
            }
        }

        public class DateTimeDeserializer : DataDeserializer {

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DateTimeDeserializer(
                Action<T, DateTime?> valueSetter,
                Func<string?, DateTime?> parser,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                DateTime? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null) : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.DefaultValue = defaultValue;
                this.Parser = parser;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception</param>
            /// <param name="dateFormat">The format is  used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DateTimeDeserializer(
                Action<T, DateTime?> valueSetter,
                string dateFormat,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                DateTime? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.DefaultValue = defaultValue;
                this.Parser = str => {

                    if (string.IsNullOrWhiteSpace(str)) return null;

                    if (DateTime.TryParseExact(str, dateFormat, null, DateTimeStyles.AssumeLocal, out var dt)) return dt;
                    else return null;
                };
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception</param>
            /// <param name="dateFormats">The formats are used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DateTimeDeserializer(
                Action<T, DateTime?> valueSetter,
                string[] dateFormats,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                DateTime? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.DefaultValue = defaultValue;
                this.Parser = str => {

                    if (string.IsNullOrWhiteSpace(str)) return null;

                    if (DateTime.TryParseExact(str, dateFormats, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out var dt)) return dt;
                    else if (DateTime.TryParse(str, out dt)) return dt;
                    else return null;
                };
            }

            public Action<Record, DateTime?> ValueSetter { get; private set; }
            public DateTime? DefaultValue { get; private set; }
            public Func<string?, DateTime?> Parser { get; private set; }

            public override void DeserializeTo(Record rec, Cell cell, ExcelDataReader<T> reader) {

                if (rec.Data == null) {
                    rec.Error = "Data is not initialized.";
                    return;
                }

                if (cell.DataType?.Value == CellValues.Date) {

                    var str = cell.CellValue?.InnerText;
                    if (string.IsNullOrWhiteSpace(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else if (DateTime.TryParse(str, out var date))
                        this.ValueSetter(rec, date);
                    else
                        this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a DateTime value.", this.DefaultValue, this.ValueSetter);
                }
                else if (cell.DataType?.Value == CellValues.Number || cell.DataType == null) {

                    var str = this.GetCellValue(cell, reader);
                    var val = this.Parser(str);
                    if (val.HasValue)
                        this.ValueSetter(rec, val.Value);
                    else if (double.TryParse(cell.CellValue?.InnerText, out var d) && (d < 2958466.0 && d > -657435.0)) // this is the valid range of OADate
                        this.ValueSetter(rec, DateTime.FromOADate(d));
                    else if (string.IsNullOrEmpty(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else
                        this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a DateTime value.", this.DefaultValue, this.ValueSetter);
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    var str = this.GetCellValue(cell, reader);
                    if (string.IsNullOrWhiteSpace(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else {
                        var val = this.Parser(str);
                        if (val.HasValue) this.ValueSetter(rec, val.Value);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a DateTime value.", this.DefaultValue, this.ValueSetter);
                    }
                }
            }
        }

        public class TimeSpanDeserializer : DataDeserializer {

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public TimeSpanDeserializer(
                Action<T, TimeSpan?> valueSetter,
                Func<string?, TimeSpan?>? parser = null,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                TimeSpan? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.DefaultValue = defaultValue;
                this.Parser = parser ?? (str => TimeSpan.TryParse(str, out var ts) ? ts : default);
            }

            public Action<Record, TimeSpan?> ValueSetter { get; private set; }
            public TimeSpan? DefaultValue { get; private set; }
            public Func<string?, TimeSpan?> Parser { get; private set; }

            public override void DeserializeTo(Record rec, Cell cell, ExcelDataReader<T> reader) {

                if (rec.Data == null) {
                    rec.Error = "Data is not initialized.";
                    return;
                }

                if (cell.DataType?.Value == CellValues.Date) {

                    var str = cell.CellValue?.InnerText;
                    if (string.IsNullOrWhiteSpace(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else if (DateTime.TryParse(str, out var date))
                        this.ValueSetter(rec, date.TimeOfDay);
                    else
                        this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to TimeSpan value.", this.DefaultValue, this.ValueSetter);
                }
                else if (cell.DataType?.Value == CellValues.Number || cell.DataType == null) {

                    var str = this.GetCellValue(cell, reader);
                    if (double.TryParse(cell.CellValue?.InnerText, out var d) && (d < 2958466.0 && d > -657435.0))
                        this.ValueSetter(rec, DateTime.FromOADate(d).TimeOfDay);
                    else if (string.IsNullOrEmpty(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else {
                        var val = this.Parser(str);
                        if (val.HasValue) this.ValueSetter(rec, val.Value);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a TimeSpan value.", this.DefaultValue, this.ValueSetter);
                    }
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    var str = this.GetCellValue(cell, reader);
                    if (string.IsNullOrWhiteSpace(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else {
                        var val = this.Parser(str);
                        if (val.HasValue) this.ValueSetter(rec, val.Value);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a TimeSpan value.", this.DefaultValue, this.ValueSetter);
                    }
                }
            }
        }

        public class BoolDeserializer : DataDeserializer {

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public BoolDeserializer(
                Action<T, bool?> valueSetter,
                Func<string?, bool?>? parser = null,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                bool? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.DefaultValue = defaultValue;
                this.Parser = parser ?? (str => bool.TryParse(str, out var ts) && ts);
            }

            public Action<Record, bool?> ValueSetter { get; private set; }
            public bool? DefaultValue { get; private set; }
            public Func<string?, bool?> Parser { get; private set; }

            public override void DeserializeTo(Record rec, Cell cell, ExcelDataReader<T> reader) {

                if (rec.Data == null) {
                    rec.Error = "Data is not initialized.";
                    return;
                }

                if (cell.DataType?.Value == CellValues.Boolean) {

                    var str = cell.CellValue?.InnerText;
                    if (string.IsNullOrWhiteSpace(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else if (bool.TryParse(str, out var b))
                        this.ValueSetter(rec, b);
                    else
                        this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to Boolean value.", this.DefaultValue, this.ValueSetter);
                }
                else if (cell.DataType?.Value == CellValues.Number || cell.DataType == null) {

                    var str = this.GetCellValue(cell, reader);
                    if (double.TryParse(cell.CellValue?.InnerText, out var d))
                        this.ValueSetter(rec, double.IsFinite(d) && d != 0);
                    else if (string.IsNullOrEmpty(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else {
                        var val = this.Parser(str);
                        if (val.HasValue) this.ValueSetter(rec, val.Value);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a Boolean value.", this.DefaultValue, this.ValueSetter);
                    }
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    var str = this.GetCellValue(cell, reader);
                    if (string.IsNullOrWhiteSpace(str))
                        this.ValueSetter(rec, this.DefaultValue);
                    else {
                        var val = this.Parser(str);
                        if (val.HasValue) this.ValueSetter(rec, val.Value);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a Boolean value.", this.DefaultValue, this.ValueSetter);
                    }
                }
            }
        }

        public class NumberDeserializer<V> : DataDeserializer where V : struct {

            private static readonly Func<string?, V?> stringParser;
            private static readonly Func<string?, V?> numericParser;

            static NumberDeserializer() {

                if (typeof(V) == typeof(byte)) stringParser = str => byte.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(sbyte)) stringParser = str => sbyte.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(short)) stringParser = str => short.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(ushort)) stringParser = str => ushort.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(int)) stringParser = str => int.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(uint)) stringParser = str => uint.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(long)) stringParser = str => long.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(ulong)) stringParser = str => ulong.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(float)) stringParser = str => float.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(double)) stringParser = str => double.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(decimal)) stringParser = str => decimal.TryParse(str, out var b) && b is V v ? v : null;
                else throw new ArgumentException($"The generic type argument type {typeof(V).Name} is not supported.");

                if (typeof(V) == typeof(byte)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToByte(b) is V v ? v : null;
                else if (typeof(V) == typeof(sbyte)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToSByte(b) is V v ? v : null;
                else if (typeof(V) == typeof(short)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToInt16(b) is V v ? v : null;
                else if (typeof(V) == typeof(ushort)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToUInt16(b) is V v ? v : null;
                else if (typeof(V) == typeof(int)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToInt32(b) is V v ? v : null;
                else if (typeof(V) == typeof(uint)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToUInt32(b) is V v ? v : null;
                else if (typeof(V) == typeof(long)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToInt64(b) is V v ? v : null;
                else if (typeof(V) == typeof(ulong)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToUInt64(b) is V v ? v : null;
                else if (typeof(V) == typeof(float)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToSingle(b) is V v ? v : null;
                else if (typeof(V) == typeof(double)) numericParser = str => double.TryParse(str, out var b) && b is V v ? v : null;
                else if (typeof(V) == typeof(decimal)) numericParser = str => double.TryParse(str, out var b) && double.IsFinite(b) && Convert.ToDecimal(b) is V v ? v : null;
                else throw new ArgumentException($"The generic type argument type {typeof(V).Name} is not supported.");
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object.
            /// Can define other logics in this setter, like converting the value from the cell to others; 
            /// or verifying the value, and if it's not valid, throw an exception.
            /// </param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <param name="equalsToNullChecker">Check the cell string value and  return true when the value should be treated as Null.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public NumberDeserializer(
                Action<T, V?> valueSetter,
                Func<string?, V?>? parser = null,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                V? defaultValue = null,
                Func<string, bool>? equalsToNullChecker = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName, equalsToNullChecker: equalsToNullChecker) {

                this.ValueSetter = this.GetValueSetterWrapper(valueSetter);
                this.Parser = parser;
                this.DefaultValue = defaultValue;
            }

            public Action<Record, V?> ValueSetter { get; private set; }
            public V? DefaultValue { get; private set; }
            public Func<string?, V?>? Parser { get; private set; }

            public override void DeserializeTo(Record rec, Cell cell, ExcelDataReader<T> reader) {

                if (rec.Data == null) {
                    rec.Error = "Data is not initialized.";
                    return;
                }

                if (cell.DataType?.Value == CellValues.Number || cell.DataType == null) {

                    var str = this.GetCellValue(cell, reader);
                    var val = numericParser(cell.CellValue?.InnerText);

                    if (val != null) {
                        this.ValueSetter(rec, val);
                    }
                    else if (string.IsNullOrEmpty(str)) {
                        this.ValueSetter(rec, this.DefaultValue);
                    }
                    else if (this.Parser != null) {
                        val = this.Parser(str);
                        if (val != null) this.ValueSetter(rec, val);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a {typeof(V).Name} value.", this.DefaultValue, this.ValueSetter);
                    }
                    else {
                        this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a {typeof(V).Name} value.", this.DefaultValue, this.ValueSetter);
                    }
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    var str = this.GetCellValue(cell, reader);
                    if (string.IsNullOrWhiteSpace(str)) {
                        this.ValueSetter(rec, this.DefaultValue);
                    }
                    else {
                        var val = (this.Parser ?? stringParser)(str);
                        if (val != null) this.ValueSetter(rec, val);
                        else this.SetErrorMsgAndWithDefaultValue(rec, $"The cell value ({str}) cannot be converted to a {typeof(V).Name} value.", this.DefaultValue, this.ValueSetter);
                    }
                }
            }
        }

        /// <summary>
        /// The area to read.
        /// </summary>
        public class DataArea {

            /// <summary>
            /// Specificed the area to read.
            /// </summary>
            /// <param name="indexOfRowBegin">read from which row. starts from 1.</param>
            /// <param name="indexOfColumnBegin">read from which column. starts from 1.</param>
            /// <param name="indexOfRowEnd">read till to the row. it's optional.</param>
            /// <param name="indexOfColumnEnd">read till to the column. it's optional.</param>
            public DataArea(uint indexOfRowBegin, uint indexOfColumnBegin, uint? indexOfRowEnd = null, uint? indexOfColumnEnd = null) {

                this.IndexOfRowBegin = indexOfRowBegin;
                this.IndexOfColumnBegin = indexOfColumnBegin;
                this.NameOfColumnBegin = ExcelUtil.GetColumnName(indexOfColumnBegin);
                this.IndexOfRowEnd = indexOfRowEnd;
                this.IndexOfColumnEnd = indexOfColumnEnd;
                this.NameOfColumnEnd = indexOfColumnEnd.HasValue ? ExcelUtil.GetColumnName(indexOfColumnEnd.Value) : null;
            }

            public DataArea(uint indexOfRowBegin, string nameOfColumnBegin, uint? indexOfRowEnd = null, string? nameOfColumnEnd = null)
                : this(indexOfRowBegin, ExcelUtil.GetColumnIndex(nameOfColumnBegin), indexOfRowEnd, nameOfColumnEnd == null ? null : ExcelUtil.GetColumnIndex(nameOfColumnEnd)) { }

            public uint IndexOfRowBegin { get; private set; }
            public uint IndexOfColumnBegin { get; private set; }
            public string NameOfColumnBegin { get; private set; }
            public uint? IndexOfRowEnd { get; set; }
            public uint? IndexOfColumnEnd { get; set; }
            public string? NameOfColumnEnd { get; private set; }

            public int CalcLength(uint endCol) {
                return (int)(this.IndexOfColumnEnd ?? endCol - this.IndexOfColumnBegin + 1);
            }
        }

        public class Record {

            public Record() { }

            public uint RowId { get; private set; }

            public T? Data { get; private set; }

            public Dictionary<string, string>? CellErrors { get; private set; }
            public string? Error { get; set; }

            public bool HasErrors {
                get {
                    return (this.CellErrors != null && this.CellErrors.Count > 0 && this.CellErrors.Values.Any(i => !i.StartsWith("Warning:"))
                    || (this.Error != null && !this.Error.StartsWith("[Warning]")));
                }
            }

            public bool HasWarnings {
                get {
                    return (this.CellErrors != null && this.CellErrors.Count > 0 && this.CellErrors.Values.Any(i => i.StartsWith("Warning:"))
                    || (this.Error != null && this.Error.StartsWith("[Warning]")));
                }
            }

            public string? GetErrorMessage() {
                if (this.HasErrors || this.HasWarnings) {
                    var error = this.Error ?? string.Empty;
                    error = error + " " + string.Join("; ", this.CellErrors?.Select(i => $"{i.Key}: {i.Value}") ?? []);
                    return error.Trim();
                }
                else {
                    return null;
                }
            }

            public void AddCellError(string? colName, string error) {

                if (colName == null) return;
                if (this.CellErrors == null) this.CellErrors = [];
                if (!this.CellErrors.TryAdd(colName, error)) this.CellErrors[colName] = error;
            }

            public Record GetReusedOrCreateNew(uint rowId, T? data, Func<T?> builder) {

                var rec = this;
                if (data == null) {
                    data = builder();
                    rec = new Record();
                }
                rec.RowId = rowId;
                rec.Data = data;
                rec.Error = null;
                rec.CellErrors?.Clear();

                return rec;
            }
        }
    }
}
