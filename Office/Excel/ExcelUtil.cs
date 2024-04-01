using System.Data;

namespace me.fengyj.CommonLib.Office.Excel {
    public static class ExcelUtil {

        public const int Max_Row_Count = 1_048_576;
        public const int Max_Column_Count = 16_384;

        /// <summary>
        /// Export the data of the tables in the dataset to spreadsheet. The sheet name is from the table's name.
        /// </summary>
        /// <param name="dataSet"></param>
        /// <param name="filePath"></param>
        public static void Export(DataSet dataSet, string filePath) {
            ExcelBuilder.BuildTo(filePath, dataSet);
        }

        /// <summary>
        /// Read row to list.
        /// </summary>
        /// <param name="filePath">spreadsheet file path</param>
        /// <param name="sheetNo">which sheet to read</param>
        /// <param name="indexOfRowBegin">read from which row. starts from 1.</param>
        /// <param name="indexOfColumnBegin">read from which column. starts from 1.</param>
        /// <param name="indexOfRowEnd">read till to the row. it's optional.</param>
        /// <param name="indexOfColumnEnd">read till to the column. it's optional.</param>
        /// <param name="reused">the object for reuse.</param>
        /// <returns>the length of the list in each row is variable, depends on the last cell which has value in the row.</returns>
        public static IEnumerable<List<string?>> Read(
            string filePath,
            uint sheetNo,
            uint indexOfRowBegin,
            uint indexOfColumnBegin,
            uint? indexOfRowEnd = null,
            uint? indexOfColumnEnd = null,
            List<string?>? reused = default) {

            var range = new ExcelDataReader<List<string?>>.DataArea(indexOfRowBegin, indexOfColumnBegin, indexOfRowEnd, indexOfColumnEnd);
            return Read(filePath, sheetNo, range, reused: reused);
        }

        /// <summary>
        /// Read row to list.
        /// </summary>
        /// <param name="filePath">spreadsheet file path</param>
        /// <param name="sheetNo">which sheet to read</param>
        /// <param name="range">the area to read</param>
        /// <param name="reused">the object for reuse.</param>
        /// <returns>the length of the list in each row is variable, depends on the last cell which has value in the row.</returns>
        public static IEnumerable<List<string?>> Read(
            string filePath,
            uint sheetNo,
            ExcelDataReader<List<string?>>.DataArea range,
            List<string?>? reused = default) {

            var recBuilder = (uint rowIdx) => new List<string?>();
            var defaultSetter = (List<string?> lst, uint colIdx, string? val) => {
                while (lst.Count <= colIdx - range.IndexOfColumnBegin) lst.Add(null);
                lst[(int)colIdx - (int)range.IndexOfColumnBegin] = val;
            };
            var cfg = new ExcelDataReader<List<string?>>.Config(sheetNo, range, recBuilder, defaultSetter: defaultSetter);

            using (var reader = new ExcelDataReader<List<string?>>(filePath)) {
                var data = reader.Read(cfg, reused);
                foreach (var item in data)
                    yield return item;
            }
        }

        private static string GetColumnName(string prefix, uint column) {
            return column < 26
                ? $"{prefix}{(char)('A' + column - 1)}"
                : GetColumnName(GetColumnName(prefix, (column - column % 26) / 26 - 1), column % 26);
        }

        public static string GetColumnName(uint column) {
            return GetColumnName(string.Empty, column);
        }

        public static string GetCellReference(uint row, uint column) => $"{GetColumnName(string.Empty, column)}{row}";

        public static string GetTableReference(uint rowStart, uint rowEnd, uint columnStart, uint columnEnd)
            => $"{GetCellReference(rowStart, columnStart)}:{GetCellReference(rowEnd, columnEnd)}";

        public static uint GetColumnIndex(string column) {
            var v = 0u;
            for (var i = 0; i < column.Length; i++)
                v = v * 26 + column[i] - 65 + 1;
            return v;
        }

        public static (uint, uint) GetCellRowAndColumnIndex(string cellRef) {

            var l = 0;
            for (; l < cellRef.Length; l++)
                if (!(cellRef[l] >= 'A'))
                    break;
            return (uint.Parse(cellRef[l..]), GetColumnIndex(cellRef[..l]));
        }
    }
}
