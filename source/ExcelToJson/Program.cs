using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Excel;
using Microsoft.CSharp;

namespace ExcelToJson
{

    sealed class Program
    {
        public static List<string> ExcelPaths = new List<string>();
        public static List<string> ExcelFileNames = new List<string>();

        [STAThread]
        static void Main(string[] args)
        {
            using (var ms = new MemoryStream())
            {
                ParseConfig(args);
                CreateDir();
                List<string> cs = new List<string>();
                foreach (string excelPath in ExcelPaths)
                {
                    cs.Add(CSExporter.GenClass(excelPath));
                }
                RuntimeAssembly.Compile(cs);
                
                foreach (string str in ExcelPaths)
                {
                    JsonExporter.GenJson(str);
                }
                
                GenCsManager();
                Console.WriteLine(" ************** Exprot Succeed ! *************** ");
            }
        }

        private static void ParseConfig(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments(args, Options.Default);
            Options.Default.Check();
            DirectoryInfo root = new DirectoryInfo(Options.Default.ExcelPath);
            FileInfo[] files = root.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                ExcelPaths.Add(file.FullName);
                ExcelFileNames.Add(Path.GetFileNameWithoutExtension(file.Name));
            }
        }

        private static void CreateDir()
        {
            Directory.CreateDirectory(Options.Default.ScriptPath);
            Directory.CreateDirectory(Options.Default.JsonPath);
        }
        
        private static void GenCsManager()
        {
            var template = File.ReadAllText(Options.Default.ScriptManagerTemplate);
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
            using (FileStream file = new FileStream(Path.Combine(Options.Default.ScriptPath, $"{managerName}.cs"), FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, Options.Default.Encoding))
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

