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

        public static bool MoveFile(string srcFile, string tarFileOrFolder, bool overwrite = false) {
            if (tarFileOrFolder.EndsWith(Path.DirectorySeparatorChar))
                return MoveFile(new FileInfo(srcFile), new DirectoryInfo(tarFileOrFolder), overwrite);
            else
                return MoveFile(new FileInfo(srcFile), new FileInfo(tarFileOrFolder), overwrite);
        }

        public static bool MoveFile(FileInfo srcFile, FileInfo tarFile, bool overwrite = false) {

            if (!srcFile.Exists) return false;
            if (tarFile.Directory == null) return false;
            if (tarFile.Exists && !overwrite) return false;

            PrepareFolder(tarFile.Directory);

            File.Move(srcFile.FullName, tarFile.FullName, overwrite);
            return true;
        }

        public static bool MoveFile(FileInfo srcFile, DirectoryInfo tarFolder, bool overwrite = false) {

            var tarFile = new FileInfo(Path.Combine(tarFolder.FullName, srcFile.Name));
            return MoveFile(srcFile, tarFile, overwrite);
        }

        public static bool CopyFile(string srcFile, string tarFileOrFolder, bool overwrite = false) {
            if (tarFileOrFolder.EndsWith(Path.DirectorySeparatorChar))
                return CopyFile(new FileInfo(srcFile), new DirectoryInfo(tarFileOrFolder), overwrite);
            else
                return CopyFile(new FileInfo(srcFile), new FileInfo(tarFileOrFolder), overwrite);
        }

        public static bool CopyFile(FileInfo srcFile, FileInfo tarFile, bool overwrite = false) {

            if (!srcFile.Exists) return false;
            if (tarFile.Directory == null) return false;
            if (tarFile.Exists && !overwrite) return false;

            PrepareFolder(tarFile.Directory);

            File.Copy(srcFile.FullName, tarFile.FullName, overwrite);
            return true;
        }

        public static bool CopyFile(FileInfo srcFile, DirectoryInfo tarFolder, bool overwrite = false) {

            var tarFile = new FileInfo(Path.Combine(tarFolder.FullName, srcFile.Name));
            return CopyFile(srcFile, tarFile, overwrite);
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
