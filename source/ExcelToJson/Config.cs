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

        [Option('h', "header", Required = true, HelpText = "number lines in sheet as header.")]
        public int HeaderRows { get; set; }

        // [Option('c', "encoding", Required = false, DefaultValue = "utf8-nobom", HelpText = "export file encoding.")]
        // public string Encoding { get; set; }

        [Option('l', "low_case", Required = true, DefaultValue = false, HelpText = "convert filed name to lowcase.")]
        public bool Lowcase { get; set; }

        [Option('a', "array", Required = false, DefaultValue = false,
            HelpText = "export as array, otherwise as dict object.")]
        public bool ExportArray { get; set; }
        
        [Option('c', "config_path", Required = true, DefaultValue = "", HelpText = "config file path")]
        public string ConfigPath { get; set; }
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
    }
}