using System;
using System.Collections;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Reflection;
using Excel;
using Newtonsoft.Json;

namespace ExcelToJson
{
	/// <summary>
	/// 将DataTable对象，转换成JSON string，并保存到文件中
	/// </summary>
	class JsonExporter
	{

		private static List<int> IdList = new List<int>();
		private static string TypeName;

		public static void GenJson(string excelPath)
		{
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

				var jsonName = Path.GetFileNameWithoutExtension(excelPath);
				TypeName = jsonName;
				var jsonPath = Path.Combine(Options.Default.ScriptPath, $"{jsonName}.json");

				IdList = new List<int>();
				//-- 转换为JSON字符串

				var content = ConvertExcel(sheet);
				using (FileStream file = new FileStream(jsonPath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
				{
					using (TextWriter writer = new StreamWriter(file, Options.Default.Encoding))
					{
						writer.Write(content);
					}
				}
			}
		}

		/// <summary>
		/// 以第一列为ID，转换成ID->Object的字典对象
		/// </summary>
		private static string ConvertExcel(DataTable sheet)
		{
			Dictionary<string, object> importData = new Dictionary<string, object>();

			int firstDataRow = Options.Default.HeaderRows - 1;
			for (int i = firstDataRow; i < sheet.Rows.Count; i++)
			{
				if (sheet.Rows[i].IsNull(0))
				{
					continue;
				}

				if (sheet.Rows[i][0].ToString() == "//END")
				{
					break;
				}

				if (sheet.Rows[i][0].ToString().StartsWith("#"))
				{
					continue;
				}

				DataRow row = sheet.Rows[i];
				string id = row[sheet.Columns[0]].ToString();
				
				if (string.IsNullOrEmpty(id))
					id = i.ToString();

				//动态编译后获取编译后的类型
				object data = ConvertRowData2Obj(sheet, row , RuntimeAssembly.RuntimeAsm.GetType(TypeName));
				
				if (data == null)
				{
					Console.WriteLine("按任意键退出 ！ ");
					Console.ReadKey();
					Environment.Exit(0);
				}

				importData[id] = data;
			}
			
			//-- convert to json string
			return JsonConvert.SerializeObject(importData, Formatting.Indented);
		}

		/// <summary>
		/// 把一行数据转换成一个对象，每一列是一个属性
		/// </summary>
		private static object ConvertRowData2Obj(DataTable sheet, DataRow row, Type type)
		{
			var obj = Activator.CreateInstance(type);
			
			if (row[0] is double)
			{
				int result = int.Parse(row[0].ToString());
				if (IdList.Contains(result))
				{
					Console.WriteLine(" 主键ID重复  " + row[0]);
					return null;
				}
				IdList.Add(result);
			}

			foreach (DataColumn column in sheet.Columns)
			{
				//反射获取每个字段
				var property = type.GetProperty(column.ToString(), BindingFlags.Public | BindingFlags.Instance);

				string valueStr = row[column].ToString();

				object value = null;

				var propertyType = property.PropertyType;	
				
				
				if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
				{
					IList list = (IList)Activator.CreateInstance(propertyType);
					if (!string.IsNullOrEmpty(valueStr))
					{
						var values = valueStr.Split('|');
						var genericType = propertyType.GenericTypeArguments[0];
						foreach (string s in values)
						{
							list.Add(Convert.ChangeType(s, genericType));
						}
					}

					value = list;
				}
				else if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Dictionary<,>))
				{
					IDictionary dic = (IDictionary)Activator.CreateInstance(propertyType);
					if (!string.IsNullOrEmpty(valueStr))
					{
						var values = valueStr.Split('|');
						var keyType = propertyType.GenericTypeArguments[0];
						var valueType = propertyType.GenericTypeArguments[1];
						foreach (var s in values)
						{
							var keyValues = s.Split(':');
							var key = Convert.ChangeType(keyValues[0], keyType);
							var val = Convert.ChangeType(keyValues[1], valueType);
							dic.Add(key, val);
						}
					}

					value = dic;
				}
				//值为空
				else if (string.IsNullOrEmpty(valueStr))
				{
					if (propertyType.IsValueType)
					{
						value = Activator.CreateInstance(propertyType);
					}
				}
				else
				{
					value = Convert.ChangeType(valueStr, propertyType);
				}

				property.SetValue(obj, value);
			}

			return obj;
		}
	}
}
