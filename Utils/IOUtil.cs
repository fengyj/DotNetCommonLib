namespace me.fengyj.CommonLib.Utils {
    public class IOUtil {

        public static bool DeleteFile(string filePath, bool prepareFolder = false) => DeleteFile(new FileInfo(filePath), prepareFolder);

        public static bool DeleteFile(FileInfo fileInfo, bool prepareFolder = false) {
            if (!fileInfo.Exists) {
                if (prepareFolder && fileInfo.Directory != null) PrepareFolder(fileInfo.Directory);
                return false;
            }
            var exists = fileInfo.Exists;
            if (exists) fileInfo.Delete();
            return exists;
        }

        public static bool PrepareFolder(string folder) => PrepareFolder(new DirectoryInfo(folder));

        public static bool PrepareFolder(DirectoryInfo dirInfo) {

            if (dirInfo.Exists) return false;
            dirInfo.Create();
            return true;
        }

        public static FileInfo GetTempFile(string? subFolder = null, string? filePrefix = null, string fileExt = ".dat") {

            var folder = Path.GetTempPath();
            if (!string.IsNullOrWhiteSpace(subFolder)) folder = Path.Combine(folder, subFolder);

            var fileName = Guid.NewGuid().ToString();
            fileName = $"{filePrefix ?? string.Empty}{fileName}{fileExt}";
            fileName = Path.Combine(folder, fileName);

            var fileInfo = new FileInfo(fileName);
            DeleteFile(fileInfo, prepareFolder: true);

            return fileInfo;
        }
    }
}
