using IWshRuntimeLibrary;
using OfficeOpenXml;
using System.Configuration;
using System.Text.RegularExpressions;
namespace ExсelBackup
{
    public class Program
    {
        private static string backupPath = ConfigurationManager.AppSettings.Get("BackupPath");
        private static string path = ConfigurationManager.AppSettings.Get("Path");
        private static string backupType = ConfigurationManager.AppSettings.Get("BackupType");
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (args.Length == 0)
            {
                args = new string[] { "" };
            }
            switch (args[0])
            {
                case "backup":
                    CreateBackup();
                    break;
                case "restore":
                    Restore();
                    break;
                default:
                    Console.WriteLine("Usage:" + '\r' + '\n' + "\"backup\" - create a backup" + '\r' + '\n' + "\"restore\" - restore a file from backup" + '\r' + '\n' + "Example:" + '\r' + '\n' + "ExcelBackup.exe restore");
                    break;
            }
        }
        private static void CreateBackup()
        {
            if (System.IO.File.Exists(path))
            {
                Directory.CreateDirectory(backupPath);
                if (Directory.GetFiles(backupPath).Length > 0)
                {
                    if (FileSearch.CountMains(backupPath) != 1)
                        Console.WriteLine("The Backup is corrupted :(");
                    else
                    {
                        string writeFile = Excel.CreateBlank(path, backupPath);
                        if (backupType == "3")
                        {
                            var files = FileSearch.FindFilesForRecover(backupPath);
                            for (int i = 0; i < files.Count; i++)
                            {
                                if (files.Count - 1 == i)
                                {
                                    Excel.CompareAndWrite(path, String.Concat(Path.GetDirectoryName(path), @"\temp"), writeFile);
                                    System.IO.File.Delete(String.Concat(Path.GetDirectoryName(path), @"\temp"));
                                }
                                else
                                {
                                    if (i == 0)
                                        Excel.Restore(i, files, @"\temp", true);
                                    else
                                        Excel.Restore(i, files, @"\temp", false);
                                }
                            }
                        }
                        else
                            Excel.CompareAndWrite(path, FileSearch.GetMainName(), writeFile);
                    }
                }
                else System.IO.File.Copy(path, String.Concat(backupPath, @"\main_", DateTime.Now.ToString("dd.MM.yyyy_HH-mm-ss"), ".xlsx"));
            }
            else Console.WriteLine("The main file was not found. Check the configuration.");
        }
        private static void EnterBackupNumber(List<string> files)
        {
            Console.Write("Enter a backup number: ");
            int BackupID = Convert.ToInt32(Console.ReadLine());
            if (BackupID < 1 || BackupID > files.Count)
            {
                Console.WriteLine("Incorrect number, try agan.");
                EnterBackupNumber(files);
            }
            else
            {
                if (backupType == "3")
                {
                    for (int i = 0; i < BackupID; i++)
                    {
                        if (i == 0)
                            Excel.Restore(i, files, @"\recovered.xlsx", true);
                        else
                            Excel.Restore(i, files, @"\recovered.xlsx", false);
                    }
                }
                else
                    Excel.Restore(BackupID - 1, files);
                Console.WriteLine("Done! File write as \"recovered.xlsx\". Press any key to exit.");
                Console.ReadKey();
            }
        }
        private static void Restore()
        {
            Directory.CreateDirectory(backupPath);
            var files = FileSearch.FindFilesForRecover(backupPath);
            Console.WriteLine("Backups:");
            for (int i = 0; i < files.Count; i++)
            {
                Console.WriteLine(i + 1 + ") " + Path.GetFileName(files[i]));
            }
            EnterBackupNumber(files);
        }
    }

    public static class Excel
    {
        private static string backupPath = ConfigurationManager.AppSettings.Get("BackupPath");
        private static string path = ConfigurationManager.AppSettings.Get("Path");
        private static int[] worksheetID = { 0, 1 };
        public static void Restore(int id, List<string> files)
        {
            Restore(id, files, @"\recovered.xlsx", true);
        }
        public static void Restore(int id, List<string> files, string filename, bool rewriteFile)
        {
            string writeFile = String.Concat(Path.GetDirectoryName(path), filename);
            if (rewriteFile)
                System.IO.File.Copy(files[0], writeFile, true);
            if (id != 0)
            {
                string removalMark = ConfigurationManager.AppSettings.Get("RemovalMark");
                // файл, с которого считываем
                FileInfo fileInfoRead = new FileInfo(files[id]);
                ExcelPackage packageRead = new ExcelPackage(fileInfoRead);
                // файл, в который пишем изменения
                FileInfo fileInfoWrite = new FileInfo(writeFile);
                ExcelPackage packageWrite = new ExcelPackage(fileInfoWrite);
                foreach (int ID in worksheetID)
                {
                    ExcelWorksheet worksheetRead = packageRead.Workbook.Worksheets[ID];
                    ExcelWorksheet worksheetWrite = packageWrite.Workbook.Worksheets[ID];
                    int rows = 0;
                    int columns = 0;
                    if (worksheetRead.Dimension is not null)
                    {
                        rows = worksheetRead.Dimension.End.Row;
                        columns = worksheetRead.Dimension.End.Column;
                    }
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= columns; j++)
                        {
                            if (worksheetRead.Cells[i, j].Value is not null || (worksheetRead.Cells[i, j] is not null && worksheetRead.Cells[i, j].Formula != ""))
                            {
                                worksheetWrite.Cells[i, j].Style.Numberformat = worksheetRead.Cells[i, j].Style.Numberformat;
                                if (worksheetRead.Cells[i, j].Formula != "")
                                    worksheetWrite.Cells[i, j].Formula = worksheetRead.Cells[i, j].Formula;
                                else
                                {
                                    if (worksheetRead.Cells[i, j].Value.ToString() == removalMark)
                                    {
                                        if (worksheetWrite.Cells[i, j].Value is not null)
                                            worksheetWrite.Cells[i, j].Value = "";
                                        if (worksheetWrite.Cells[i, j].Formula is not null)
                                            worksheetWrite.Cells[i, j].Formula = "";
                                    }
                                    else
                                    {
                                        worksheetWrite.Cells[i, j].Value = worksheetRead.Cells[i, j].Value;
                                    }
                                }
                            }
                        }
                    }
                }
                packageWrite.Save();
            }
        }
        public static string CreateBlank(string Path, string BackupPath)
        {
            string file = String.Concat(BackupPath, @"\", FileSearch.FindLastId(BackupPath) + 1, " ", DateTime.Now.ToString("dd.MM.yyyy_HH-mm-ss"), ".xlsx");
            System.IO.File.Copy(Path, file);
            FileInfo fileInfo = new FileInfo(file);
            ExcelPackage package = new ExcelPackage(fileInfo);
            foreach (int ID in worksheetID)
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[ID];
                int rows = 0;
                int columns = 0;
                if (worksheet.Dimension is not null)
                {
                    rows = worksheet.Dimension.End.Row;
                    columns = worksheet.Dimension.End.Column;
                }
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {
                        if (worksheet.Cells[i, j] is not null)
                        {
                            if (worksheet.Cells[i, j].Value is not null)
                                worksheet.Cells[i, j].Value = "";
                            if (worksheet.Cells[i, j].Formula is not null)
                                worksheet.Cells[i, j].Formula = "";
                        }
                    }
                }
            }
            package.Save();
            return file;
        }

        public static void CompareAndWrite(string sourceFile, string compareFile, string writeFile)
        {
            string removalMark = ConfigurationManager.AppSettings.Get("RemovalMark");
            // файл, с которого считываем
            FileInfo fileInfoRead = new FileInfo(sourceFile);
            ExcelPackage packageRead = new ExcelPackage(fileInfoRead);
            // файл, с которым сравниваем
            FileInfo fileInfoCompare = new FileInfo(compareFile);
            ExcelPackage packageCompare = new ExcelPackage(fileInfoCompare);
            // файл, в который пишем изменения
            FileInfo fileInfoWrite = new FileInfo(writeFile);
            ExcelPackage packageWrite = new ExcelPackage(fileInfoWrite);
            foreach (int ID in worksheetID)
            {
                ExcelWorksheet worksheetRead = packageRead.Workbook.Worksheets[ID];
                ExcelWorksheet worksheetCompare = packageCompare.Workbook.Worksheets[ID];
                ExcelWorksheet worksheetWrite = packageWrite.Workbook.Worksheets[ID];
                int rows = 0;
                int columns = 0;
                if (worksheetRead.Dimension is not null)
                {
                    rows = worksheetRead.Dimension.End.Row;
                    columns = worksheetRead.Dimension.End.Column;
                }
                if (worksheetCompare.Dimension is not null)
                {
                    if (worksheetRead.Dimension is not null)
                    {
                        if (worksheetCompare.Dimension.End.Row > worksheetRead.Dimension.End.Row)
                            rows = worksheetCompare.Dimension.End.Row;
                        if (worksheetCompare.Dimension.End.Column > worksheetRead.Dimension.End.Column)
                            columns = worksheetCompare.Dimension.End.Column;
                    }
                    else
                    {
                        rows = worksheetCompare.Dimension.End.Row;
                        columns = worksheetCompare.Dimension.End.Column;
                    }
                }

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {
                        if (worksheetRead.Cells[i, j].Value is not null || (worksheetRead.Cells[i, j] is not null && worksheetRead.Cells[i, j].Formula != "")) //сейчас в ячейке что-то есть
                        {
                            worksheetWrite.Cells[i, j].Style.Numberformat = worksheetRead.Cells[i, j].Style.Numberformat; // копируем тип данных
                            if (worksheetCompare.Cells[i, j].Value is not null || (worksheetCompare.Cells[i, j] is not null && worksheetCompare.Cells[i, j].Formula != "")) //раньше в ячейке что-то было
                            {
                                if (worksheetRead.Cells[i, j].Formula != "") //сейчас в ячейке формула
                                {
                                    if (worksheetCompare.Cells[i, j].Formula != "") //если формула была и есть
                                    {
                                        //Если формулы различаются
                                        if (worksheetRead.Cells[i, j].Formula.ToString() != worksheetCompare.Cells[i, j].Formula.ToString())
                                            worksheetWrite.Cells[i, j].Formula = worksheetRead.Cells[i, j].Formula;
                                        //Если формулы одинаковые, ничего не делаем
                                    }
                                    else //если сейчас формула есть, а раньше не было
                                    {
                                        worksheetWrite.Cells[i, j].Formula = worksheetRead.Cells[i, j].Formula;
                                    }
                                }
                                else // в ячейке и сейчас и раньше что-то было, но сейчас в ней текст
                                {
                                    if (worksheetCompare.Cells[i, j].Formula != "") //раньше была формула, сейчас текст
                                    {
                                        worksheetWrite.Cells[i, j].Value = worksheetRead.Cells[i, j].Value;
                                    }
                                    else //раньше был текст, сейчас тоже текст
                                    {
                                        //Если текст различается

                                        if (worksheetRead.Cells[i, j].Value.ToString() != worksheetCompare.Cells[i, j].Value.ToString())
                                            worksheetWrite.Cells[i, j].Value = worksheetRead.Cells[i, j].Value;
                                        //Если текст одинаковый, ничего не делаем
                                    }
                                }
                            }
                            else // сейчас в ячейке что-то есть, раньше в ней ничего не было
                            {
                                if (worksheetRead.Cells[i, j].Formula != "")
                                    worksheetWrite.Cells[i, j].Formula = worksheetRead.Cells[i, j].Formula;
                                else
                                    worksheetWrite.Cells[i, j].Value = worksheetRead.Cells[i, j].Value;
                            }

                        }
                        else // сейчас в ячейке пусто
                        {
                            if (worksheetCompare.Cells[i, j].Value is not null || (worksheetCompare.Cells[i, j] is not null && worksheetCompare.Cells[i, j].Formula != "")) //раньше в ячейке что-то было
                            {
                                worksheetWrite.Cells[i, j].Value = removalMark;
                            }
                        }
                    }
                }
            }
            packageWrite.Save();
        }
    }
    public static class FileSearch
    {
        public static string GetFullPath(string file)
        {
            if (System.IO.File.Exists(file))
            {
                WshShell shell = new WshShell();
                IWshShortcut link = (IWshShortcut)shell.CreateShortcut(file);
                return link.TargetPath;
            }
            else return "";
        }
        private static string MainName { get; set; }
        public static string GetMainName()
        {
            return MainName;
        }
        public static List<string> FindFilesForRecover(string path)
        {
            var sortedFiles = new List<string>();
            Regex pattern = new Regex(@"\S+[\\](\d+\s|main_)\d{2}.\d{2}.\d{4}_\d{2}-\d{2}-\d{2}.xlsx");
            string[] files = Directory.GetFiles(path);
            foreach (string file in files)
            {
                MatchCollection matches = pattern.Matches(file);
                if (matches.Count > 0)
                {
                    if (matches[0].Groups[1].Value == "main_")
                        sortedFiles.Insert(0, file);
                    else
                        sortedFiles.Add(file);
                }
            }
            return sortedFiles;
        }
        public static int FindLastId(string path)
        {
            int ID = 0;
            Regex regex = new Regex(@"\S+[\\](\d+)\s\d{2}.\d{2}.\d{4}_\d{2}-\d{2}-\d{2}.xlsx");
            string[] files = Directory.GetFiles(path);
            foreach (string file in files)
            {
                MatchCollection matches = regex.Matches(file);
                if (matches.Count > 0)
                {
                    if (int.Parse(matches[0].Groups[1].Value) > ID)
                        ID = int.Parse(matches[0].Groups[1].Value);
                }
            }
            return ID;
        }
        public static int CountMains(string path)
        {
            string pattern = @"(\S+[\\]main_\d{2}.\d{2}.\d{4}_\d{2}-\d{2}-\d{2}.xlsx)";
            int flag = 0;
            string[] files = Directory.GetFiles(path);
            foreach (string file in files)
            {
                if (Regex.IsMatch(file, pattern/*, RegexOptions.IgnoreCase*/))
                {
                    MainName = file;
                    flag++;
                }
            }
            return flag;
        }
    }
}
