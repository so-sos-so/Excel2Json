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
            if (args.Length <= 0)
            {
                Console.WriteLine("must provide parameter \"-c config_file_path\"");
                return;
            }
            CommandLine.Parser.Default.ParseArguments(args, Config.Options);
            Config.Options.ConfigPath = Path.GetFullPath(Config.Options.ConfigPath);
            if (File.Exists(Config.Options.ConfigPath))
            {
                if (!Directory.Exists(Config.Options.ExcelPath))
                {
                    Console.WriteLine($"{Config.Options.ExcelPath} is not exist, can not continue, Check it !");
                    Console.ReadKey();
                    return;
                }
                
                DirectoryInfo root = new DirectoryInfo(Config.Options.ExcelPath);
                FileInfo[] files = root.GetFiles("*.xlsx");
                foreach (FileInfo file in files)
                {
                    ExcelPaths.Add(file.FullName);
                    ExcelFileNames.Add(Path.GetFileNameWithoutExtension(file.Name));
                }
            }
            else
            {
                Console.WriteLine($"{Config.Options} not exist, set it ");
                Console.ReadKey();
            }
        }

        private static void loadExcel(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            //-- start import
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

        private static Dictionary<string, string> ParseConfigFile(string _path)
        {
            // 读取文件的源路径及其读取流
            StreamReader srReadFile = new StreamReader(_path);
            Dictionary<string, string> pathDic = new Dictionary<string, string>();
            while (!srReadFile.EndOfStream)
            {
                string str = srReadFile.ReadLine();
                if (string.IsNullOrEmpty(str))
                    continue;
                var match = Regex.Match(str, @"(\w+)=(.+)");
                pathDic.Add(match.Groups[1].Value, match.Groups[2].Value);
            }
            srReadFile.Close();
            return pathDic;
        }

        private static void GenCsManager()
        {
            var template = File.ReadAllText(Config.Options.ScriptManagerTemplate);
            string managerName = Regex.Match(template, @"class\s+(\w+)").Groups[1].Value;
            StringBuilder field_dics = new StringBuilder();
            StringBuilder load_config = new StringBuilder();
            StringBuilder get_by_id = new StringBuilder();
            foreach (var fileName in ExcelFileNames)
            {
                field_dics.AppendLine(
                    $"\tpublic static Dictionary<int, {fileName}> {fileName} = new Dictionary<int, {fileName}>();");
                load_config.AppendLine($"\t\t{fileName} = LoadOneConfig<{fileName}>(\"{fileName}\");");
                get_by_id.Append(GetByIdTemplate(fileName));
            }
            template = template.Replace("#field_dics", field_dics.ToString());
            template = template.Replace("#load_config", load_config.ToString());
            template = template.Replace("#get_by_id", get_by_id.ToString());
            template = template.Replace("#json_path", Config.Options.JsonLoadPath);
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
            return sb;
        }

    }
}

