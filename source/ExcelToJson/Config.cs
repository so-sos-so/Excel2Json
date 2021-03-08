using System;
using System.IO;
using System.Reflection;
using System.Text;
using CommandLine;

namespace ExcelToJson
{

    /// <summary>
    /// 命令行参数定义
    /// </summary>
    public sealed class Options
    {

        public static Options Default = new Options();

        public Options()
        {
            HeaderRows = 3;
        }

        [Option("header")] public int HeaderRows { get; set; }

        [Option("encoding", Required = false, DefaultValue = "UTF8", HelpText = "export file encoding.")]
        public string EncodingStr { get; set; }

        public Encoding Encoding =>
            (Encoding) typeof(Encoding).GetProperty(EncodingStr, BindingFlags.Static | BindingFlags.Public)
                .GetValue(null);

        [Option( "excel_path")]
        public string ExcelPath { get; set; }
        [Option( "json_path")]
        public string JsonPath { get; set; }
        [Option( "script_path")]
        public string ScriptPath { get; set; }
        [Option( "script_template_path")]
        public string ScriptTemplate { get; set; }
        [Option( "script_manager_template_path")]
        public string ScriptManagerTemplate { get; set; }

        public void Check()
        {
            if(!Directory.Exists(ExcelPath))
                Console.WriteLine($"{ExcelPath} not exist");
            if(!File.Exists(ScriptTemplate))
                Console.WriteLine($"{ScriptTemplate} not exist");
            if(!File.Exists(ScriptManagerTemplate))
                Console.WriteLine($"{ScriptManagerTemplate} not exist");
        }
    }
}