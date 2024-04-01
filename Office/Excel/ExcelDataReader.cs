using System.Globalization;

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
            this.doc = SpreadsheetDocument.Open(filePath, false);
        }

        public IEnumerable<T> Read(
            Config config,
            T? reused = default) {

            var sheet = this.doc.WorkbookPart?.Workbook.Sheets?.Skip((int)config.SheetNo - 1).FirstOrDefault();
            if (sheet == null || sheet is not Sheet s) yield break;
            var sheetPart = this.doc.WorkbookPart?.GetPartById(s.Id?.Value ?? "");
            if (sheetPart == null) yield break;

            // todo: get col index if it's not provided

#pragma warning disable CS8629 // Nullable value type may be null.
            var dict = config.Deserializers.Where(i => i.ColumnIndex.HasValue).GroupBy(i => i.ColumnIndex.Value).ToDictionary(i => i.Key, i => i.ToList());
#pragma warning restore CS8629 // Nullable value type may be null.
            var isReadViaTableColumns = config.Deserializers.Count != dict.Count; // if not equal, means there are some deserializers don't provide the columnIndex.

            using (var reader = OpenXmlReader.Create(sheetPart)) {

                var currentRowIdx = 0u;
                var currentColIdx = 0u;

                T? rec = default;

                while (reader.Read()) {

                    if (reader.ElementType == typeof(Row)) {

                        if (reader.IsStartElement) {

                            var rowId = GetRowId(reader);
                            if (rowId != null) {
                                for (var i = currentRowIdx + 1; i < rowId; i++) {

                                    currentRowIdx++;
                                    currentColIdx = 0; // reset col idx, because it's a new row now.

                                    // return an empty record.
                                    // because user cannot distinglish the row in the sheet is not existed nor an empty row,
                                    // to avoid the confusion, always return an empty record.
                                    var rowCanReturn = isReadViaTableColumns ? (currentRowIdx > config.Area.IndexOfRowBegin) : (currentRowIdx >= config.Area.IndexOfRowBegin);
                                    if (rowCanReturn) {

                                        rec = reused ?? config.RecordBuilder(currentRowIdx);
                                        yield return rec;
                                    }
                                }
                            }

                            currentRowIdx++;
                            currentColIdx = 0; // reset col idx, because it's a new row now.
                            if (!isReadViaTableColumns)
                                rec = reused ?? config.RecordBuilder(currentRowIdx);
                        }
                        else {
                            // check if needs to skip the header row
                            var rowCanReturn = isReadViaTableColumns ? (currentRowIdx > config.Area.IndexOfRowBegin) : (currentRowIdx >= config.Area.IndexOfRowBegin);
                            if (rowCanReturn && currentColIdx >= config.Area.IndexOfColumnBegin)
                                if (rec != null)
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
                                else if (config.DefaultSetter != null) {
                                    config.DefaultSetter(rec, currentColIdx, this.GetStringCellValue(EmptyCell));
                                }
                            }
                        }

                        currentColIdx++;

                        if (currentRowIdx < config.Area.IndexOfRowBegin || currentColIdx < config.Area.IndexOfColumnBegin) continue;
                        if (config.Area.IndexOfColumnEnd.HasValue && currentColIdx > config.Area.IndexOfColumnEnd.Value) continue;

                        if (isReadViaTableColumns && currentRowIdx == config.Area.IndexOfRowBegin) {
                            var name = cell.CellValue?.InnerText;
                            var matched = config.Deserializers.Where(d => d.ColumnName == name && !d.ColumnIndex.HasValue);
                            foreach (var d in matched) {
                                d.ColumnIndex = currentColIdx;
                                d.ColumnName = ExcelUtil.GetColumnName(currentColIdx);
                                if (dict.TryGetValue(currentColIdx, out var lst)) lst.Add(d);
                                else dict.Add(currentColIdx, new List<DataDeserializer> { d });
                            }
                        }
                        else if (dict.TryGetValue(currentColIdx, out var deserializers)) {
                            deserializers.ForEach(d => d.DeserializeTo(rec, cell, this));
                        }
                        else if (config.DefaultSetter != null) {
                            config.DefaultSetter(rec, currentColIdx, this.GetStringCellValue(cell));
                        }
                    }
                }

                config.Area.IndexOfRowEnd = currentRowIdx;
            }
        }

        public string? GetStringCellValue(Cell cell) {

            if (cell.DataType?.Value == CellValues.SharedString) {
                return int.TryParse(cell.CellValue?.InnerText, out var sid) ? this.GetSharedString(sid) : null;
            }
            else {
                return cell == EmptyCell ? null : cell.CellValue?.InnerText;
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
            /// <param name="area">Define the range of the cells to read. the end or row and/or the end of column can be omitted.</param>
            /// <param name="recBuilder">The builder function to create an object resprents the row in the sheet. The arg is the current row index.</param>
            /// <param name="deserializers">Define the functions how to convert cell value to the values in the object.</param>
            /// <param name="defaultSetter">If necessary, when deserializers don't contain the function to handle the cell value for a particular column, can set a default one.
            /// The second arg is the current col index.
            /// </param>
            public Config(uint sheetNo, DataArea area, Func<uint, T> recBuilder, List<DataDeserializer>? deserializers = null, Action<T, uint, string?>? defaultSetter = null) {

                this.SheetNo = sheetNo;
                this.Area = area;
                this.Deserializers = deserializers ?? [];
                this.RecordBuilder = recBuilder;
                this.DefaultSetter = defaultSetter;
            }

            /// <summary>
            /// The index of the sheet contains the data to read. Starts from 1.
            /// </summary>
            public uint SheetNo { get; private set; }
            public DataArea Area { get; private set; }
            public List<DataDeserializer> Deserializers { get; private set; }
            public Func<uint, T> RecordBuilder { get; private set; }
            public Action<T, uint, string?>? DefaultSetter { get; private set; }

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
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DataDeserializer(uint? colIdx = null, string? colName = null, string? tblColName = null) {

                if (colIdx == null && colName == null && tblColName == null)
                    throw new ArgumentException("At least one of the arguments should have value.");

                if (colIdx.HasValue) this.ColumnIndex = colIdx.Value;
                if (colName != null) this.ColumnName = colName;
                this.TableColumnName = tblColName;
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

            public abstract void DeserializeTo(T data, Cell cell, ExcelDataReader<T> reader);
        }

        public class TextDeserializer : DataDeserializer {

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public TextDeserializer(Action<T, string?> valueSetter, uint? colIdx = null, string? colName = null, string? tblColName = null, string? defaultValue = null)
                : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
                this.DefaultValue = defaultValue;
            }

            public Action<T, string?> ValueSetter { get; private set; }
            public string? DefaultValue { get; private set; }
            public override void DeserializeTo(T data, Cell cell, ExcelDataReader<T> reader) {

                var val = reader.GetStringCellValue(cell);
                this.ValueSetter(data, val ?? this.DefaultValue);
            }
        }

        public class DateTimeDeserializer : DataDeserializer {

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DateTimeDeserializer(
                Action<T, DateTime?> valueSetter,
                Func<string?, DateTime?> parser,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                DateTime? defaultValue = null) : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
                this.DefaultValue = defaultValue;
                this.Parser = parser;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="dateFormat">The format is  used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DateTimeDeserializer(
                Action<T, DateTime?> valueSetter,
                string dateFormat,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                DateTime? defaultValue = null) : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
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
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="dateFormats">The formats are used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public DateTimeDeserializer(
                Action<T, DateTime?> valueSetter,
                string[] dateFormats,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                DateTime? defaultValue = null) : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
                this.DefaultValue = defaultValue;
                this.Parser = str => {

                    if (string.IsNullOrWhiteSpace(str)) return null;

                    if (DateTime.TryParseExact(str, dateFormats, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out var dt)) return dt;
                    else if (DateTime.TryParse(str, out dt)) return dt;
                    else return null;
                };
            }

            public Action<T, DateTime?> ValueSetter { get; private set; }
            public DateTime? DefaultValue { get; private set; }
            public Func<string?, DateTime?> Parser { get; private set; }

            public override void DeserializeTo(T data, Cell cell, ExcelDataReader<T> reader) {

                if (cell.DataType?.Value == CellValues.Date) {

                    if (DateTime.TryParse(cell.CellValue?.InnerText, out var date))
                        this.ValueSetter(data, date);
                    else
                        this.ValueSetter(data, this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.Number || cell.DataType == null) {

                    if (double.TryParse(cell.CellValue?.InnerText, out var d))
                        this.ValueSetter(data, DateTime.FromOADate(d));
                    else
                        this.ValueSetter(data, this.Parser(reader.GetStringCellValue(cell)) ?? this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    this.ValueSetter(data, this.Parser(reader.GetStringCellValue(cell)) ?? this.DefaultValue);
                }
            }
        }

        public class TimeSpanDeserializer : DataDeserializer {

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public TimeSpanDeserializer(
                Action<T, TimeSpan?> valueSetter,
                Func<string?, TimeSpan?>? parser = null,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                TimeSpan? defaultValue = null) : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
                this.DefaultValue = defaultValue;
                this.Parser = parser ?? (str => TimeSpan.TryParse(str, out var ts) ? ts : default);
            }

            public Action<T, TimeSpan?> ValueSetter { get; private set; }
            public TimeSpan? DefaultValue { get; private set; }
            public Func<string?, TimeSpan?> Parser { get; private set; }

            public override void DeserializeTo(T data, Cell cell, ExcelDataReader<T> reader) {

                if (cell.DataType?.Value == CellValues.Date) {

                    if (DateTime.TryParse(cell.CellValue?.InnerText, out var date))
                        this.ValueSetter(data, date.TimeOfDay);
                    else
                        this.ValueSetter(data, this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.Number || cell.DataType == null) {

                    if (double.TryParse(cell.CellValue?.InnerText, out var d))
                        this.ValueSetter(data, DateTime.FromOADate(d).TimeOfDay);
                    else
                        this.ValueSetter(data, this.Parser(reader.GetStringCellValue(cell)) ?? this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    this.ValueSetter(data, this.Parser(reader.GetStringCellValue(cell)) ?? this.DefaultValue);
                }
            }
        }

        public class BoolDeserializer : DataDeserializer {

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public BoolDeserializer(
                Action<T, bool?> valueSetter,
                Func<string?, bool?>? parser = null,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                bool? defaultValue = null) : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
                this.DefaultValue = defaultValue;
                this.Parser = parser ?? (str => bool.TryParse(str, out var ts) ? ts : default);
            }

            public Action<T, bool?> ValueSetter { get; private set; }
            public bool? DefaultValue { get; private set; }
            public Func<string?, bool?> Parser { get; private set; }

            public override void DeserializeTo(T data, Cell cell, ExcelDataReader<T> reader) {

                if (cell.DataType?.Value == CellValues.Boolean) {

                    if (bool.TryParse(cell.CellValue?.InnerText, out var b))
                        this.ValueSetter(data, b);
                    else
                        this.ValueSetter(data, this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.Number) {

                    if (double.TryParse(cell.CellValue?.InnerText, out var d))
                        this.ValueSetter(data, d != 0);
                    else
                        this.ValueSetter(data, this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString) {

                    this.ValueSetter(data, this.Parser(reader.GetStringCellValue(cell)) ?? this.DefaultValue);
                }
            }
        }

        public class NumberDeserializer<V> : DataDeserializer {

            private static readonly Func<string?, V?> defaultParser;

            static NumberDeserializer() {

                if (typeof(V) == typeof(byte)) defaultParser = str => byte.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(sbyte)) defaultParser = str => sbyte.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(short)) defaultParser = str => short.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(ushort)) defaultParser = str => ushort.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(int)) defaultParser = str => int.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(uint)) defaultParser = str => uint.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(long)) defaultParser = str => long.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(ulong)) defaultParser = str => ulong.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(float)) defaultParser = str => float.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(double)) defaultParser = str => double.TryParse(str, out var b) && b is V v ? v : default(V);
                else if (typeof(V) == typeof(decimal)) defaultParser = str => decimal.TryParse(str, out var b) && b is V v ? v : default(V);
                else throw new ArgumentException($"The generic type argument type {typeof(V).Name} is not supported.");
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="valueSetter">The action to update the value to the object</param>
            /// <param name="parser">The parser is used when the cell data type is string.</param>
            /// <param name="colIdx">The value in which column in the sheet. Index starts from 1.</param>
            /// <param name="colName">The value in which column in the sheet. The name starts from A.</param>
            /// <param name="tblColName">The value in which table column in the sheet. If specificed this, means the first row in the range is the table header, will be skipped</param>
            /// <param name="defaultValue">The default value when the cell doesn't have value nor cannot be parsed.</param>
            /// <exception cref="ArgumentException">colIdx, colName, tblColName, one of them must be specificed.</exception>
            public NumberDeserializer(
                Action<T, V?> valueSetter,
                Func<string?, V?>? parser = null,
                uint? colIdx = null,
                string? colName = null,
                string? tblColName = null,
                V? defaultValue = default) : base(colIdx: colIdx, colName: colName, tblColName: tblColName) {

                this.ValueSetter = valueSetter;
                this.Parser = parser;
                this.DefaultValue = defaultValue;
            }

            public Action<T, V?> ValueSetter { get; private set; }
            public V? DefaultValue { get; private set; }
            public Func<string?, V?>? Parser { get; private set; }

            public override void DeserializeTo(T data, Cell cell, ExcelDataReader<T> reader) {

                if (cell.DataType?.Value == CellValues.Number) {
                    this.ValueSetter(data, defaultParser(cell.CellValue?.InnerText) ?? this.DefaultValue);
                }
                else if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString
                    || cell.DataType?.Value == CellValues.SharedString || cell.DataType == null) {
                    this.ValueSetter(data, (this.Parser ?? defaultParser)(reader.GetStringCellValue(cell)) ?? this.DefaultValue);
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
    }
}
