using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using Excel;

namespace ExcelToJson
{
	/// <summary>
	/// 根据表头，生成C#类定义数据结构
	/// 表头使用三行定义：字段名称、字段类型、注释
	/// </summary>
	class CSExporter
	{
		
		struct FieldDef
		{
			public string Name;
			public string Type;
			public string Remark;
		}

		public static string GenClass(string excelPath)
		{
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
				
				var csName = Path.GetFileNameWithoutExtension(excelPath);
				var scriptPath = Path.Combine(Options.Default.ScriptPath, $"{csName}.cs");
				var fieldList = ParseField(sheet);
				var template = File.ReadAllText(Options.Default.ScriptTemplate);
				template = template.Replace("#class_name", csName);

				StringBuilder sb = new StringBuilder();
				foreach (FieldDef field in fieldList)
				{
					sb.AppendFormat("\tpublic {0} {1} {{ get; set; }}// {2}", field.Type, field.Name, field.Remark);
					sb.AppendLine();
				}
				
				template = template.Replace("#fields", sb.ToString());
				using (FileStream file = new FileStream(scriptPath, FileMode.Create, FileAccess.Write))
				{
					using (TextWriter writer = new StreamWriter(file, Options.Default.Encoding))
						writer.Write(template);
				}
				Console.WriteLine(" -- " + csName);
				return template;
			}
		}
		
		private const string ArrayPre = "array_";
		private const string DicPre = "dic_";

		private static List<FieldDef> ParseField(DataTable sheet)
		{
			var result = new List<FieldDef>();
			DataRow typeRow = sheet.Rows[0];
			DataRow commentRow = sheet.Rows[1];

			foreach (DataColumn column in sheet.Columns)
			{
				FieldDef field;
				field.Name = column.ToString();
				var typeStr = typeRow[column].ToString();
				if (typeStr.Contains(ArrayPre))
				{
                    
					var listTypeStr = typeStr.Replace(ArrayPre, "");
					typeStr = $"List<{listTypeStr}>";
				}
				else if(typeStr.Contains(DicPre))
				{
                    
					var match = Regex.Match(typeStr, @"\w+_(\w+)_(\w+)");
					typeStr = $"Dictionary<{match.Groups[1].Value}, {match.Groups[2].Value}>";
				}

				field.Type = typeStr;
				field.Remark = Regex.Replace(commentRow[column].ToString(), "[\r\n\t]", " ", RegexOptions.Compiled); ;
				result.Add(field);
			}

			return result;
		}
	}
}
