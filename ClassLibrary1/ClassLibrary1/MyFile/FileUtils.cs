using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using YanBo;

namespace YanBo_CG
{
    public class FileUtils
    {
        public static string ExcelFilter = "Excel files(*.xls)|*.xls;*.xlsx";

        public static string OpenDir(string title)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = title;
            string result;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                result = folderBrowserDialog.SelectedPath;
            }
            else
            {
                result = null;
            }
            return result;
        }

        internal static string SelectSingleExcelFile()
        {
            return FileUtils.SelectSingleFile("Excel files(*.xls)|*.xls;*.xlsx");
        }

        internal IList<string> FileInforToNames(IList<FileInfo> fileInfos)
        {
            IList<string> list = new List<string>();
            foreach (FileInfo current in fileInfos)
            {
                list.Add(current.Name);
            }
            return list;
        }

        internal IList<string> FileInforToDirs(IList<FileInfo> fileInfos)
        {
            IList<string> list = new List<string>();
            foreach (FileInfo current in fileInfos)
            {
                list.Add(current.DirectoryName);
            }
            return list;
        }

        internal static IList<string> SelectSingleExcelFiles()
        {
            return FileUtils.SelectFiles("Excel files(*.xls)|*.xls;*.xlsx");
        }

        internal IList<string> FileInforTofullNames(IList<FileInfo> fileInfos)
        {
            IList<string> list = new List<string>();
            foreach (FileInfo current in fileInfos)
            {
                list.Add(current.FullName);
            }
            return list;
        }

        internal static IList<FileInfo> PathToFileInfo(IList<string> pathList)
        {
            IList<FileInfo> list = new List<FileInfo>();
            foreach (string current in pathList)
            {
                if (File.Exists(current))
                {
                    list.Add(new FileInfo(current));
                }
            }
            return list;
        }

        public static bool CopyFile(string fileName, string saveFileName, bool cover)
        {
            bool result;
            if (!cover && File.Exists(saveFileName))
            {
                MessageBox.Show("文件已经存在：" + saveFileName);
                result = false;
            }
            else
            {
                if (File.Exists(fileName))
                {
                    string text = saveFileName.Substring(0, saveFileName.LastIndexOf("\\"));
                    if (!Directory.Exists(text))
                    {
                        try
                        {
                            Directory.CreateDirectory(text);
                        }
                        catch
                        {
                            throw new Exception("路径不合法：" + text);
                        }
                    }
                    try
                    {
                        File.Copy(fileName, saveFileName, true);
                        result = true;
                        return result;
                    }
                    catch
                    {
                        throw new Exception("文件名不合法：" + text);
                    }
                }
                result = false;
            }
            return result;
        }

        public static bool CopyFileCover(string fileName, string saveFileName)
        {
            string directoryName = Path.GetDirectoryName(saveFileName);
            if (!Directory.Exists(directoryName))
            {
                try
                {
                    Directory.CreateDirectory(directoryName);
                }
                catch
                {
                    throw new Exception("路径不合法：" + directoryName);
                }
            }
            bool result;
            try
            {
                File.Copy(fileName, saveFileName, true);
                result = true;
            }
            catch
            {
                throw new Exception("文件名不合法：" + directoryName);
            }
            return result;
        }

        public static bool CopyFileNoCover(string fileName, string saveFileName)
        {
            return !File.Exists(saveFileName) && FileUtils.CopyFileCover(fileName, saveFileName);
        }

        internal void RomveExcept(IList<string> list, IList<string> filters, FileSelectRelation relation)
        {
            int num = list.Count;
            for (int i = 0; i < num; i++)
            {
                string fileName = list[i];
                if (!this.CheckFile(fileName, filters, relation))
                {
                    list.RemoveAt(i);
                    i--;
                    num--;
                }
            }
        }

        private bool CheckFile(string fileName, IList<string> filters, FileSelectRelation relation)
        {
            int num = fileName.LastIndexOf("\\") + 1;
            fileName = fileName.Substring(num, fileName.Length - num);
            bool result;
            foreach (string current in filters)
            {
                if (this.CheckFile(fileName, current, relation))
                {
                    result = true;
                    return result;
                }
            }
            result = false;
            return result;
        }

        private bool CheckFile(string compareName, string filterName, FileSelectRelation relation)
        {
            bool result;
            switch (relation)
            {
                case FileSelectRelation.StartsWith:
                    if (compareName.StartsWith(filterName))
                    {
                        result = true;
                        return result;
                    }
                    break;
                case FileSelectRelation.EndsWith:
                    if (compareName.EndsWith(filterName))
                    {
                        result = true;
                        return result;
                    }
                    break;
                case FileSelectRelation.Contains:
                    if (compareName.Contains(filterName))
                    {
                        result = true;
                        return result;
                    }
                    break;
                case FileSelectRelation.Equals:
                    if (compareName.Equals(filterName))
                    {
                        result = true;
                        return result;
                    }
                    break;
                case FileSelectRelation.All:
                    result = true;
                    return result;
            }
            result = false;
            return result;
        }

        internal IList<string> SeleFileDir(string dir, bool level)
        {
            IList<string> list = new List<string>();
            return FileUtils.SelectAfterDirFiles(dir, list, level);
        }

        public static IList<string> SelectAfterDirFiles(string dirName, IList<string> list, bool level)
        {
            string[] files = Directory.GetFiles(dirName);
            string[] array = files;
            for (int i = 0; i < array.Length; i++)
            {
                string item = array[i];
                list.Add(item);
            }
            if (!level)
            {
                string[] directories = Directory.GetDirectories(dirName);
                array = directories;
                for (int i = 0; i < array.Length; i++)
                {
                    string text = array[i];
                    if (Directory.Exists(text))
                    {
                        FileUtils.SelectAfterDirFiles(text, list, level);
                    }
                }
            }
            return list;
        }

        public static IList<string> SelectDirs(string dir)
        {
            IList<string> list = new List<string>();
            FileUtils.SelectDirs(dir, list);
            return list;
        }

        private static void SelectDirs(string dir, IList<string> list)
        {
            string[] directories = Directory.GetDirectories(dir);
            string[] array = directories;
            for (int i = 0; i < array.Length; i++)
            {
                string text = array[i];
                list.Add(text);
                FileUtils.SelectDirs(text, list);
            }
        }

        public static IList<string> SelectAfterDirFiles(string filterName, string dirName, IList<string> list, FileSelectRelation relation, bool level)
        {
            string[] files = Directory.GetFiles(dirName);
            string[] array = files;
            for (int i = 0; i < array.Length; i++)
            {
                string text = array[i];
                string text2 = text.Substring(text.LastIndexOf("\\")).Replace("\\", "");
                switch (relation)
                {
                    case FileSelectRelation.StartsWith:
                        if (text2.StartsWith(filterName))
                        {
                            list.Add(text);
                        }
                        break;
                    case FileSelectRelation.EndsWith:
                        if (text2.EndsWith(filterName))
                        {
                            list.Add(text);
                        }
                        break;
                    case FileSelectRelation.Contains:
                        if (text2.Contains(filterName))
                        {
                            list.Add(text);
                        }
                        break;
                    case FileSelectRelation.Equals:
                        if (text2.Equals(filterName))
                        {
                            list.Add(text);
                        }
                        break;
                    case FileSelectRelation.All:
                        list.Add(text);
                        break;
                }
            }
            if (!level)
            {
                string[] directories = Directory.GetDirectories(dirName);
                array = directories;
                for (int i = 0; i < array.Length; i++)
                {
                    string text3 = array[i];
                    if (Directory.Exists(text3))
                    {
                        FileUtils.SelectAfterDirFiles(filterName, text3, list, relation, level);
                    }
                }
            }
            return list;
        }

        public static List<string> SelectFiles(string filter, string title)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = filter;
            openFileDialog.Title = title;
            openFileDialog.ValidateNames = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.CheckFileExists = true;
            List<string> list = new List<string>();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] fileNames = openFileDialog.FileNames;
                for (int i = 0; i < fileNames.Length; i++)
                {
                    string item = fileNames[i];
                    list.Add(item);
                }
            }
            return list;
        }

        public static List<string> SelectFiles(string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = filter;
            openFileDialog.ValidateNames = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.CheckFileExists = true;
            List<string> list = new List<string>();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] fileNames = openFileDialog.FileNames;
                for (int i = 0; i < fileNames.Length; i++)
                {
                    string item = fileNames[i];
                    list.Add(item);
                }
            }
            return list;
        }

        public static string SelectSingleFile(string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = filter;
            openFileDialog.ValidateNames = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.CheckFileExists = true;
            string result = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                result = openFileDialog.FileName;
            }
            return result;
        }

        public static string SelectSingleFile(string title, string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = title;
            openFileDialog.Filter = filter;
            openFileDialog.ValidateNames = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.CheckFileExists = true;
            string result = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                result = openFileDialog.FileName;
            }
            return result;
        }

        public static string SeleFileDir()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            return folderBrowserDialog.SelectedPath;
        }

        public static string SeleFileDir(string title)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = title;
            folderBrowserDialog.ShowDialog();
            return folderBrowserDialog.SelectedPath;
        }

        public static string SeleFileDir(string title, string dirRoot)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = title;
            folderBrowserDialog.SelectedPath = dirRoot;
            folderBrowserDialog.ShowDialog();
            return folderBrowserDialog.SelectedPath;
        }

        public static string[] GetFileNameArray(string fileName)
        {
            string[] array = new string[2];
            array[0] = fileName.Substring(0, fileName.LastIndexOf("\\"));
            array[1] = fileName.Substring(array[0].Length, fileName.Length - array[0].Length);
            return array;
        }

        public static string SaveFile(string filter)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = "C:\\ZJDTemplete";
            saveFileDialog.Filter = filter;
            saveFileDialog.RestoreDirectory = true;
            DialogResult dialogResult = saveFileDialog.ShowDialog();
            string result;
            if (dialogResult == DialogResult.OK && saveFileDialog.FileName.Length > 0)
            {
                result = saveFileDialog.FileName;
            }
            else
            {
                result = null;
            }
            return result;
        }

        public static void CopyDirectory(string srcPath, string destPath)
        {
            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileSystemInfos = directoryInfo.GetFileSystemInfos();
                if (!Directory.Exists(destPath))
                {
                    Directory.CreateDirectory(destPath);
                }
                FileSystemInfo[] array = fileSystemInfos;
                for (int i = 0; i < array.Length; i++)
                {
                    FileSystemInfo fileSystemInfo = array[i];
                    if (fileSystemInfo is DirectoryInfo)
                    {
                        if (!Directory.Exists(destPath + "\\" + fileSystemInfo.Name))
                        {
                            Directory.CreateDirectory(destPath + "\\" + fileSystemInfo.Name);
                        }
                        FileUtils.CopyDirectory(fileSystemInfo.FullName, destPath + "\\" + fileSystemInfo.Name);
                    }
                    else
                    {
                        File.Copy(fileSystemInfo.FullName, destPath + "\\" + fileSystemInfo.Name, true);
                    }
                }
            }
            catch (Exception var_3_DB)
            {
                throw;
            }
        }

        internal static Dictionary<string, IList<string>> GetDirDic(IList<string> fileList)
        {
            Dictionary<string, IList<string>> dictionary = new Dictionary<string, IList<string>>();
            foreach (string current in fileList)
            {
                string directoryName = Path.GetDirectoryName(current);
                IList<string> list;
                if (dictionary.TryGetValue(directoryName, out list))
                {
                    list.Add(current);
                }
                else
                {
                    list = new List<string>();
                    list.Add(current);
                    dictionary.Add(directoryName, list);
                }
            }
            return dictionary;
        }

        public static string FindContStrDir(IList<string> dirList, string str)
        {
            string result;
            foreach (string current in dirList)
            {
                if (current.Contains(str))
                {
                    result = current;
                    return result;
                }
            }
            result = null;
            return result;
        }

        public static string FindContStrPath(IList<string> list, string str1, string str2)
        {
            string result;
            if (Utils.IsStrNull(str2))
            {
                foreach (string current in list)
                {
                    if (current.Contains(str1))
                    {
                        result = current;
                        return result;
                    }
                }
            }
            else
            {
                foreach (string current in list)
                {
                    if (current.Contains(str1) && current.Contains(str2))
                    {
                        result = current;
                        return result;
                    }
                }
            }
            result = null;
            return result;
        }

        internal static string SelectSingleExcelFile(string title)
        {
            return FileUtils.SelectSingleFile(title, "Excel files(*.xls)|*.xls;*.xlsx");
        }

        public static string SaveFileName()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            string result;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;
                result = fileName;
            }
            else
            {
                result = null;
            }
            return result;
        }

        internal static string SaveFileName(string saveName, string filter)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = filter;
            saveFileDialog.FileName = saveName;
            saveFileDialog.RestoreDirectory = true;
            DialogResult dialogResult = saveFileDialog.ShowDialog();
            string result;
            if (dialogResult == DialogResult.OK && saveFileDialog.FileName.Length > 0)
            {
                result = saveFileDialog.FileName;
            }
            else
            {
                result = null;
            }
            return result;
        }

        public static string SaveFileName(string dir, string saveName, string filter)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = dir;
            saveFileDialog.Filter = filter;
            saveFileDialog.FileName = saveName;
            saveFileDialog.RestoreDirectory = true;
            DialogResult dialogResult = saveFileDialog.ShowDialog();
            string result;
            if (dialogResult == DialogResult.OK && saveFileDialog.FileName.Length > 0)
            {
                result = saveFileDialog.FileName;
            }
            else
            {
                result = null;
            }
            return result;
        }

        internal static string SaveFileName(string dir)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = dir;
            saveFileDialog.RestoreDirectory = true;
            DialogResult dialogResult = saveFileDialog.ShowDialog();
            string result;
            if (dialogResult == DialogResult.OK && saveFileDialog.FileName.Length > 0)
            {
                result = saveFileDialog.FileName;
            }
            else
            {
                result = null;
            }
            return result;
        }
    }
}
