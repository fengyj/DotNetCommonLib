using System.Data;

namespace me.fengyj.CommonLib.Utils.Data {
    public class DataTableUtil {

        /// <summary>
        /// Create DataTable via the data.
        /// </summary>
        /// <param name="values">values can be null or empty only when headers is not null.</param>
        /// <param name="headers">it's length should match the length in the values. if it's null, will be generated as Column1, Column2, ...</param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static DataTable CreateTable(object?[][]? values, string[]? headers = null, string? tableName = null) {

            if (headers == null && values != null && values.Length > 0) {
                headers = new string[values[0].Length];
                for (var i = 0; i < headers.Length; i++) headers[i] = $"Column{i}";
            }

            if (headers == null || headers.Length == 0)
                throw new ArgumentException($"The {nameof(values)} cannot be null nor empty when {nameof(headers)} is null.");

            var types = new Type[headers.Length];
            if (values != null) {
                foreach (var row in values) {
                    for (var i = 0; i < types.Length; i++) {
                        var v = row[i];
                        if (v == null) continue;
                        var t = v.GetType();
                        if (types[i] == null) types[i] = t;
                        else if (types[i] != t) types[i] = typeof(string);
                    }
                }
            }
            for (var i = 0; i < types.Length; i++)
                if (types[i] == null) types[i] = typeof(string);

            var table = new DataTable();
            if (tableName != null) table.TableName = tableName;

            for (var i = 0; i < headers.Length; i++) {
                table.Columns.Add(new DataColumn(headers[i], types[i]));
            }

            if (values != null) {
                foreach (var item in values) {

                    var row = table.NewRow();
                    for (var i = 0; i < headers.Length; i++) {

                        var v = item[i];
                        if (v != null) {
                            if (v.GetType() == types[i]) row[i] = v;
                            else row[i] = v.ToString();
                        }
                    }
                    table.Rows.Add(row);
                }
            }

            return table;
        }
    }
}
