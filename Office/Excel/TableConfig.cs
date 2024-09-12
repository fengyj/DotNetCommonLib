﻿using System.Text.RegularExpressions;

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

        private static readonly Regex CustomFormatRegex = new Regex("(?<value>(\\{0(:(?<format>([^\\{\\}]+)))?\\}))", RegexOptions.Compiled);

        public static readonly ColumnTotalFunction None = new ColumnTotalFunction(0, "None");
        public static readonly ColumnTotalFunction Average = new ColumnTotalFunction(101, "Average");
        public static readonly ColumnTotalFunction CountNumbers = new ColumnTotalFunction(102, "Count Numbers");
        public static readonly ColumnTotalFunction Count = new ColumnTotalFunction(103, "Count");
        public static readonly ColumnTotalFunction Max = new ColumnTotalFunction(104, "Max");
        public static readonly ColumnTotalFunction Min = new ColumnTotalFunction(105, "Min");
        public static readonly ColumnTotalFunction StdDev = new ColumnTotalFunction(107, "StdDev");
        public static readonly ColumnTotalFunction Sum = new ColumnTotalFunction(109, "Sum");
        public static readonly ColumnTotalFunction Var = new ColumnTotalFunction(110, "Var");

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
                CustomFormat = format
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
        /// SUBTOTAL(109,[Col A]) / SUBTOTAL(3,[Col B])
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

        public override ColumnTotalFunction WithCustomFormat(string format) => throw new NotSupportedException();
    }
}
