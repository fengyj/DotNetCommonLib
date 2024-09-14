using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Office.Excel {
    public class TableConfig<T> {

        public TableConfig(List<ITableColumnConfig<T>> columns, string? tableName = null, TableStyle? style = null) {
            this.Columns = columns;
            this.TableName = tableName;
            this.TableStyle = style ?? new();
        }

        public string? TableName { get; private set; }
        public List<ITableColumnConfig<T>> Columns { get; private set; }
        public TableStyle TableStyle { get; private set; }
    }

    public interface ITableColumnConfig<in T> {

        public string ColumnName { get; set; }
        public CellStyle? CellStyle { get; set; }
        public Type? DataType { get; set; }
        public bool HasDataGetter { get; }
        public object? GetDataObject(T t);
        public ColumnTotalFunction TotalFunction { get; set; }
    }

    public class TableColumnConfig<T, C> : ITableColumnConfig<T> {

        public TableColumnConfig(
            string colName,
            CellStyle? style = null,
            Type? dataType = null,
            Func<T, C?>? dataGetter = null,
            ColumnTotalFunction? totalFunction = null) {

            this.ColumnName = colName.Trim();
            this.CellStyle = style;
            this.DataType = dataType ?? dataGetter?.GetType().GenericTypeArguments[1];
            if (this.DataType == typeof(Nullable<>))
                this.DataType = this.DataType.GenericTypeArguments[0];
            this.DataGetter = dataGetter;
            this.TotalFunction = totalFunction ?? ColumnTotalFunction.None;
        }

        public string ColumnName { get; set; }
        public CellStyle? CellStyle { get; set; }
        public Type? DataType { get; set; }
        public Func<T, C?>? DataGetter { get; private set; }
        public ColumnTotalFunction TotalFunction { get; set; }

        public bool HasDataGetter => this.DataGetter != null;

        public object? GetDataObject(T t) {
            if (this.DataGetter == null || t == null) return default;
            return this.DataGetter(t);
        }

        public C? GetData(T t) {
            if (this.DataGetter == null || t == null) return default;
            return this.DataGetter(t);
        }
    }

    public class ColumnTotalFunction {

        private static readonly Regex CustomFormatRegex = new("(?<value>(\\{0(:(?<format>([^\\{\\}]+)))?\\}))", RegexOptions.Compiled);

        public static readonly ColumnTotalFunction None = new(0, "None");
        public static readonly ColumnTotalFunction Average = new(101, "Average");
        public static readonly ColumnTotalFunction CountNumbers = new(102, "Count Numbers");
        public static readonly ColumnTotalFunction Count = new(103, "Count");
        public static readonly ColumnTotalFunction Max = new(104, "Max");
        public static readonly ColumnTotalFunction Min = new(105, "Min");
        public static readonly ColumnTotalFunction Product = new(106, "Product");
        public static readonly ColumnTotalFunction StdDev = new(107, "StdDev");
        public static readonly ColumnTotalFunction StdDevP = new(108, "StdDevP");
        public static readonly ColumnTotalFunction Sum = new(109, "Sum");
        public static readonly ColumnTotalFunction Var = new(110, "Var");
        public static readonly ColumnTotalFunction VarP = new(111, "VarP");

        protected ColumnTotalFunction(uint code, string name) {
            this.FunctionCode = code;
            this.FunctionName = name;
        }

        public virtual uint FunctionCode { get; private set; }
        public virtual string FunctionName { get; private set; }
        /// <summary>
        /// Define the content of cell. {0} is the placeholder of the total value. If it's not specificed, the default content will be FunctionName concat the total value, like Count: 10.
        /// </summary>
        /// <example>Total Price: {0:$0.00}</example>
        public virtual string? CustomFormat { get; private set; }
        public virtual bool IncludeHiddenRows { get; private set; } = false;

        /// <summary>
        /// Get the formula for generating Excel
        /// </summary>
        /// <param name="tableName">The table name defined in the Excel</param>
        /// <param name="columnName">The column name of the table.</param>
        /// <param name="format">The cell format defined on the column.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public virtual string GetFormula(string tableName, string columnName, string? format) {

            var valueFormula = $"SUBTOTAL({this.FunctionCode}, {tableName}[{columnName}])";
            if (string.IsNullOrWhiteSpace(this.CustomFormat))
                this.CustomFormat = $"{this.FunctionName}: {{0}}";

            if (this.FunctionCode % 100 == 2 || this.FunctionCode % 100 == 3)
                format = DefaultStyleConfig.Numbering.DefaultInteger.FormatString; // for count functions, the format shouldn't inherited from column's

            var matchResult = CustomFormatRegex.Match(this.CustomFormat);
            if (!matchResult.Success) throw new ArgumentException($"The {nameof(this.CustomFormat)} {this.CustomFormat} is not valid.");

            var groupValue = matchResult.Groups["value"];
            if (!groupValue.Success) throw new ArgumentException($"The {nameof(this.CustomFormat)} {this.CustomFormat} is not valid.");

            var groupFormat = matchResult.Groups["format"];
            var strFormat = groupFormat.Success ? groupFormat.Value : format;
            if (!string.IsNullOrWhiteSpace(strFormat)) valueFormula = $"TEXT({valueFormula}, \"{strFormat}\")";

            var part1 = groupValue.Index == 0 ? string.Empty : this.CustomFormat.Substring(0, groupValue.Index);
            var part2 = valueFormula;
            var part3 = groupValue.Index + groupValue.Length == this.CustomFormat.Length ? string.Empty : this.CustomFormat.Substring(groupValue.Index + groupValue.Length);
            return $"CONCATENATE(\"{part1}\", {part2}, \"{part3}\")";
        }

        /// <summary>
        /// The format must contain a placeholder ({0}), otherwise it's an invalid format, then will use the default format (defined on column) for output.
        /// </summary>
        /// <example>Total Price: {0:$0.00}</example>
        /// <remarks>If it's not specified, the format will be ;FuncName]: [Value]</remarks>
        /// <param name="format"></param>
        /// <returns></returns>
        public virtual ColumnTotalFunction WithCustomFormat(string format) {
            var func = new ColumnTotalFunction(this.FunctionCode, this.FunctionName) {
                CustomFormat = format,
                IncludeHiddenRows = this.IncludeHiddenRows
            };
            return func;
        }

        /// <summary>
        /// Set the formula using the hidden rows or now.
        /// </summary>
        /// <param name="hiddenRows"></param>
        /// <returns></returns>
        public virtual ColumnTotalFunction WithHiddenRows(bool hiddenRows) {

            if (this.FunctionCode == 0) return this;

            var code = this.FunctionCode;
            if (code < 100 && hiddenRows) code += 100;
            else if (code > 100 && !hiddenRows) code -= 100;
            else return this;

            var func = new ColumnTotalFunction(code, this.FunctionName) {
                CustomFormat = this.CustomFormat,
                IncludeHiddenRows = !hiddenRows
            };
            return func;
        }
    }

    /// <summary>
    /// If the functions defined in <see cref="ColumnTotalFunction"/> is not enough, can define a customized formula.
    /// </summary>
    public class CustomColumnTotalFunction : ColumnTotalFunction {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="formula">The content of the Excel formula (No leading '=').</param>
        public CustomColumnTotalFunction(string formula) : base(0, string.Empty) {
            this.Formula = formula;
        }

        /// <summary>
        /// The content of the Excel formula (No leading '=').
        /// </summary>
        /// <example>
        /// like <code>SUBTOTAL(109,[Col A]) / SUBTOTAL(3,[Col B])</code> or even more complicated one 
        /// <code>SUMPRODUCT(SUBTOTAL(3, OFFSET(Table1[Col A], ROW(Table1[Col A])-MIN(ROW(Table1[Col A])), 0, 1)), --(Table1[Col A]="Yes"))</code>
        /// </example>
        public string Formula { get; private set; }

        /// <summary>
        /// Return the customized Excel formula.
        /// </summary>
        /// <param name="tableName">useless</param>
        /// <param name="columnName">useless</param>
        /// <param name="format">useless</param>
        /// <returns></returns>
        public override string GetFormula(string tableName, string columnName, string? format) {

            return this.Formula;
        }

        public override uint FunctionCode => throw new NotSupportedException();

        public override string FunctionName => throw new NotSupportedException();

        public override string? CustomFormat => throw new NotSupportedException();

        public override bool IncludeHiddenRows => throw new NotSupportedException();

        public override ColumnTotalFunction WithCustomFormat(string format) => throw new NotSupportedException();

        public override ColumnTotalFunction WithHiddenRows(bool hiddenRows) => throw new NotSupportedException();
    }
}
