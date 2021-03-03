using System;
using System.IO;
using System.Text;
using CommandLine;

namespace ExcelToJson
{
    public static class Config
    {
        public static Encoding Encoding = new UTF8Encoding(false);
        public static Options Options = new Options();
    }
    
    /// <summary>
    /// 命令行参数定义
    /// </summary>
    public sealed class Options
    {
        public Options()
        {
            HeaderRows = 3;
            //Encoding = "utf8-nobom";
            Lowcase = false;
            ExportArray = false;
        }

        public Options(Options options)
        {
            HeaderRows = options.HeaderRows;
            //Encoding = options.Encoding;
            Lowcase = options.Lowcase;
            ExportArray = options.Lowcase;
        }

        [Option("header")]
        public int HeaderRows { get; set; }

        // [Option('c', "encoding", Required = false, DefaultValue = "utf8-nobom", HelpText = "export file encoding.")]
        // public string Encoding { get; set; }

        [Option("low_case")]
        public bool Lowcase { get; set; }

        [Option("array")]
        public bool ExportArray { get; set; }
        
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
        [Option( "json_load")]
        public string JsonLoadPath { get; set; }

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