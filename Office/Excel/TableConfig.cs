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

        public void Verify() {

            if (this.Columns == null || this.Columns.Count == 0)
                throw new InvalidDataException($"The {nameof(this.Columns)} cannot be null nor empty.");

            foreach (var i in this.Columns)
                if (string.IsNullOrWhiteSpace(i.ColumnName))
                    throw new InvalidDataException($"The {nameof(i.ColumnName)} of {nameof(this.Columns)} cannot be null nor empty.");

            var groupByName = this.Columns.GroupBy(i => i.ColumnName);
            foreach (var g in groupByName)
                if (g.Count() > 1)
                    throw new InvalidDataException($"The name {g.Key} is duplicated.");
        }
    }

    public interface ITableColumnConfig<T> {

        public string ColumnName { get; set; }
        public CellStyle? CellStyle { get; set; }
        public Type? DataType { get; set; }
        public bool HasDataGetter { get; }
        public object? GetDataObject(T t);
    }

    public class TableColumnConfig<T, C> : ITableColumnConfig<T> {

        public TableColumnConfig(
            string colName,
            CellStyle? style = null,
            Type? dataType = null,
            Func<T, C?>? dataGetter = null) {

            this.ColumnName = colName.Trim();
            this.CellStyle = style;
            this.DataType = dataType ?? dataGetter?.GetType().GenericTypeArguments[1];
            if (this.DataType == typeof(Nullable<>))
                this.DataType = this.DataType.GenericTypeArguments[0];
            this.DataGetter = dataGetter;
        }

        public string ColumnName { get; set; }
        public CellStyle? CellStyle { get; set; }
        public Type? DataType { get; set; }
        public Func<T, C?>? DataGetter { get; private set; }

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
}
