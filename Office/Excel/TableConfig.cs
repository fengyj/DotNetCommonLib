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

    public interface ITableColumnConfig<T> {

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
            if (this.DataGetter == null || t == null) return default(C);
            return this.DataGetter(t);
        }

        public C? GetData(T t) {
            if (this.DataGetter == null || t == null) return default(C);
            return this.DataGetter(t);
        }
    }

    public class ColumnTotalFunction {

        public static ColumnTotalFunction None = new ColumnTotalFunction(0, "None");
        public static ColumnTotalFunction Average = new ColumnTotalFunction(101, "Average");
        public static ColumnTotalFunction CountNumbers = new ColumnTotalFunction(102, "Count Numbers");
        public static ColumnTotalFunction Count = new ColumnTotalFunction(103, "Count");
        public static ColumnTotalFunction Max = new ColumnTotalFunction(104, "Max");
        public static ColumnTotalFunction Min = new ColumnTotalFunction(105, "Min");
        public static ColumnTotalFunction StdDev = new ColumnTotalFunction(107, "StdDev");
        public static ColumnTotalFunction Sum = new ColumnTotalFunction(109, "Sum");
        public static ColumnTotalFunction Var = new ColumnTotalFunction(110, "Var");

        private ColumnTotalFunction(uint code, string name) {
            this.FunctionCode = code;
            this.FunctionName = name;
        }

        public uint FunctionCode { get; private set; }
        public string FunctionName { get; private set; }
        /// <summary>
        /// Define the content of cell. {0} is the placeholder of the total value. If it's not specificed, the default content will be FunctionName concat the total value, like Count: 10.
        /// </summary>
        public string? CustomFormat { get; private set; }

        public string GetFormula(string tableName, string columnName, string? format) {

            var valueFormula = $"SUBTOTAL({this.FunctionCode}, {tableName}[{columnName}])";
            if (!string.IsNullOrWhiteSpace(format)) valueFormula = $"TEXT({valueFormula}, \"{format}\")";
            if (this.CustomFormat == null) {
                return $"CONCATENATE(\"{this.FunctionName}: \", {valueFormula})"; // todo: there is an issue that the format will be lost when using concatenate to join with other texts. find a solution to fix it.
            }
            else {
                var idx = this.CustomFormat.IndexOf("{0}");
                if (idx < 0)
                    return $"CONCATENATE(\"{this.FunctionName}: \", {valueFormula})";
                var parameters = new List<string>();
                if (idx > 0) parameters.Add($"\"{this.CustomFormat[..idx]}\"");
                parameters.Add(valueFormula);
                if (idx + 3 < this.CustomFormat.Length) parameters.Add($"\"{this.CustomFormat[(idx + 3)..]}\"");
                return $"CONCATENATE({string.Join(", ", parameters)})";
            }
        }

        /// <summary>
        /// The format must contain a placeholder ({0}), otherwise it's an invalid format, then will use the default format for output.
        /// </summary>
        /// <param name="format"></param>
        /// <returns></returns>
        public ColumnTotalFunction WithCustomFormat(string format) {
            var func = new ColumnTotalFunction(this.FunctionCode, this.FunctionName);
            func.CustomFormat = format;
            return func;
        }
    }
}
