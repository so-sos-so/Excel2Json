using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
			public string name;
			public string type;
			public string comment;
		}

		private static List<FieldDef> _fieldList;

		public static void GenClass(string path, DataTable sheet)
		{
			//-- First Row as Column Name
			if (sheet.Rows.Count < 2)
				return;

			ParseField(sheet);
			var csName = Path.GetFileNameWithoutExtension(path);
			var template = File.ReadAllText(Config.Options.ScriptTemplate);
			template = template.Replace("#class_name", csName);

			StringBuilder fieldsb = new StringBuilder();
			foreach (FieldDef field in _fieldList)
			{
				fieldsb.AppendFormat("\tpublic {0} {1} {{ get; set; }}// {2}", field.type, field.name, field.comment);
				fieldsb.AppendLine();
			}
			template = template.Replace("#fields", fieldsb.ToString());
			using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write))
			{
				using (TextWriter writer = new StreamWriter(file, Config.Encoding))
					writer.Write(template);
			}
		}

		private static void ParseField(DataTable sheet)
		{
			_fieldList = new List<FieldDef>();
			DataRow typeRow = sheet.Rows[0];
			DataRow commentRow = sheet.Rows[1];

			foreach (DataColumn column in sheet.Columns)
			{
				FieldDef field;
				field.name = column.ToString();
				field.type = typeRow[column].ToString();
				if (field.type == "arraystring")
				{
					field.type = "List<string>";
				}
				else if (field.type == "arrayint")
				{
					field.type = "List<int>";
				}
				else if (field.type == "arrayfloat")
				{
					field.type = "List<float>";
				}
				else if (field.type == "array")
				{
					field.type = "List<string>";
				}
				else if (field.type == "dic")
				{
					field.type = "Dictionary<int, int>";
				}

				field.comment = Regex.Replace(commentRow[column].ToString(), "[\r\n\t]", " ", RegexOptions.Compiled); ;

				_fieldList.Add(field);
			}
		}
	}
}
