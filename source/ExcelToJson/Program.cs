using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Excel;

namespace ExcelToJson
{

    sealed class Program
    {
        public static List<string> ExcelPaths = new List<string>();
        public static List<string> ExcelFileNames = new List<string>();

        [STAThread]
        static void Main(string[] args)
        {
            ParseConfig(args);
            foreach (string str in ExcelPaths)
            {
                loadExcel(str);
            }
            GenCsManager();
            Console.WriteLine(" ************** Exprot Succeed ! *************** ");
        }

        private static void ParseConfig(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments(args, Config.Options);
            Config.Options.Check();
            DirectoryInfo root = new DirectoryInfo(Config.Options.ExcelPath);
            FileInfo[] files = root.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                ExcelPaths.Add(file.FullName);
                ExcelFileNames.Add(Path.GetFileNameWithoutExtension(file.Name));
            }
        }

        private static void loadExcel(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            Directory.CreateDirectory(Config.Options.ScriptPath);
            Directory.CreateDirectory(Config.Options.JsonPath);
            loadExcel(path, Path.Combine(Config.Options.ScriptPath, $"{name}.cs"),
                Path.Combine(Config.Options.JsonPath, $"{name}.json"));
        }

        private static void loadExcel(string excelPath, string scriptPath, string jsonPath)
        {
            string excelName = Path.GetFileNameWithoutExtension(excelPath);

            // 加载Excel文件
            using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

                excelReader.IsFirstRowAsColumnNames = true;
                DataSet book = excelReader.AsDataSet();

                // 数据检测
                if (book.Tables.Count < 1)
                {
                    throw new Exception("Excel file is empty: " + excelPath);
                }

                // 取得数据
                DataTable sheet = book.Tables[0];
                if (sheet.Rows.Count <= 0)
                {
                    throw new Exception("Excel Sheet is empty: " + excelPath);
                }
                
                CSExporter.GenClass(scriptPath, sheet);
                JsonExporter.GenJson(jsonPath, sheet);
                Console.WriteLine(" -- " + excelName + ".xslx");
            }
        }

        private static void GenCsManager()
        {
            var template = File.ReadAllText(Config.Options.ScriptManagerTemplate);
            string managerName = Regex.Match(template, @"class\s+(\w+)").Groups[1].Value;
            StringBuilder field_dics = new StringBuilder();
            StringBuilder get_by_id = new StringBuilder();
            foreach (var fileName in ExcelFileNames)
            {
                field_dics.AppendLine(
                    $"\tpublic static Dictionary<int, {fileName}> {fileName} {{ get; private set; }}");
                get_by_id.Append(GetByIdTemplate(fileName));
            }
            template = template.Replace("#field_dics", field_dics.ToString());
            template = template.Replace("#get_by_id", get_by_id.ToString());
            using (FileStream file = new FileStream(Path.Combine(Config.Options.ScriptPath, $"{managerName}.cs"), FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, Config.Encoding))
                    writer.Write(template);
            }
        }

        private static StringBuilder GetByIdTemplate(string fieldName)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"\tpublic static {fieldName} Get{fieldName}ById(int id)");
            sb.AppendLine("\t{");
            sb.AppendLine($"\t\tif({fieldName}.TryGetValue(id, out var data))");
            sb.AppendLine("\t\t\treturn data;");
            sb.AppendLine("\t\treturn null;");
            sb.AppendLine("\t}");
            sb.AppendLine();
            return sb;
        }
    }
}

