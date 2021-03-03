using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ExcelToJson
{
	/// <summary>
	/// 将DataTable对象，转换成JSON string，并保存到文件中
	/// </summary>
	class JsonExporter
	{

		private static List<int> IdList = new List<int>();

		public static void GenJson(string path, DataTable sheet)
		{
			if (sheet.Columns.Count <= 0)
				return;
			if (sheet.Rows.Count <= 0)
				return;

			IdList = new List<int>();
			//-- 转换为JSON字符串

			var content = convertDict(sheet);  //字典
			using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
			{
				using (TextWriter writer = new StreamWriter(file, Config.Encoding))
				{
					writer.Write(content);
				}
			}
		}

		/// <summary>
		/// 以第一列为ID，转换成ID->Object的字典对象
		/// </summary>
		private static string convertDict(DataTable sheet)
		{
			Dictionary<string, object> importData = new Dictionary<string, object>();

			int firstDataRow = Config.Options.HeaderRows - 1;
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
				string ID = row[sheet.Columns[0]].ToString();

				if (ID.Length <= 0)
					ID = string.Format("row_{0}", i);
				object data = convertRowData(sheet, row, Config.Options.Lowcase, firstDataRow);
				if (data == null)
				{
					Console.WriteLine("按任意键退出 ！ ");
					Console.ReadKey();
					Environment.Exit(0);
				}

				importData[ID] = data;
			}

			//-- convert to json string
			return JsonConvert.SerializeObject(importData, Formatting.Indented);
		}

		private static bool IsNumeric(string str) //接收一个string类型的参数,保存到str里
		{
			if (str == null || str.Length == 0)    //验证这个参数是否为空
				return false;                           //是，就返回False
			ASCIIEncoding ascii = new ASCIIEncoding();//new ASCIIEncoding 的实例
			byte[] bytestr = ascii.GetBytes(str);         //把string类型的参数保存到数组里

			foreach (byte c in bytestr)                   //遍历这个数组里的内容
			{
				if (c < 48 || c > 57)                          //判断是否为数字
				{
					return false;                              //不是，就返回False
				}
			}
			return true;                                        //是，就返回True
		}

		/// <summary>
		/// 把一行数据转换成一个对象，每一列是一个属性
		/// </summary>
		private static object convertRowData(DataTable sheet, DataRow row, bool lowcase, int firstDataRow)
		{
			var rowData = new Dictionary<string, object>();
			int col = 0;

			//Console.WriteLine("sheet  " + sheet.ToString());

			//Console.WriteLine("row  " + row.ToString());

			if (row[0] is double)
			{
				int result = int.Parse(row[0].ToString());
				if (IdList.Contains(result))
				{
					Console.WriteLine(" 主键ID重复  " + row[0] + "  阻断性BUG！");
					return null;
				}
				else
				{
					IdList.Add(result);
				}
			}

			int index = 0;
			foreach (DataColumn column in sheet.Columns)
			{
				object value = row[column];

				/*-------------------- 类型检测-数据为空 【数组】------------------------*/
				string rowType = sheet.Rows[0][column].ToString();

				//string typeName = 
				if (value.ToString() == "")
				{
					if (rowType == "array")
					{
						value = new List<int>();
					}
					else if (rowType == "arrayint")
					{
						value = new List<int>();
					}
					else if (rowType == "arraystring")
					{
						value = new List<string>();
					}
					else if (rowType == "arrayfloat")
					{
						value = new List<float>();
					}
				}
				else if (rowType == "arraystring" || rowType == "arrayint" || rowType == "array" || rowType == "arrayfloat")
				{
					/*-------------------- 类型检测-数据不空 【数组】------------------------*/
					if (value.ToString().Contains("|"))
					{
						string[] strlist = value.ToString().Split('|');
						//智能识别数字和文字
						if (IsNumeric(strlist[0]))
						{
							List<int> content = new List<int>();
							foreach (string str in strlist)
							{
								string temp = str;
								if (int.TryParse(temp, out int tag))
								{
									content.Add(int.Parse(temp));
								}
							}
							value = content;

						}
						else
						{
							List<string> content = new List<string>();
							foreach (string str in strlist)
							{
								string temp = str;
								content.Add(temp.ToString());
							}
							value = content;
						}
					}
					else if (rowType == "array")
					{
						List<int> content = new List<int>();
						string temp = value.ToString().Trim();
						if (int.TryParse(temp, out int tag))
						{
							content.Add(int.Parse(temp));
						}
						value = content;
					}
					else if (rowType == "arrayint")
					{
						List<int> content = new List<int>();
						string temp = value.ToString().Trim();
						if (int.TryParse(temp, out int tag))
						{
							content.Add(int.Parse(temp));
						}
						value = content;
					}
					else if (rowType == "arrayfloat")
					{
						List<float> content = new List<float>();
						string temp = value.ToString().Trim();
						if (float.TryParse(temp, out float tag))
						{
							content.Add(float.Parse(temp));
						}
						value = content;
					}
					else if (rowType == "arraystring")
					{
						List<string> content = new List<string>();
						content.Add(value.ToString());
						value = content;
					}

				}

				//字典
				//if (value.ToString().StartsWith("{") && value.ToString().EndsWith("}"))
				//{
				//    Dictionary<string, int> dic = new Dictionary<string, int>();
				//    string[] strlist = value.ToString().Split('}');
				//    foreach (string str in strlist)
				//    {
				//        string temp = str;
				//        if (temp.StartsWith("{"))
				//        {
				//            temp = temp.Substring(1);
				//            string[] valuelist = temp.Split(',');
				//            try
				//            {
				//                dic.Add(valuelist[0], int.Parse(valuelist[1]));
				//            }
				//            catch
				//            {
				//                Console.WriteLine("Error: Have wrong type in dictionary ");
				//                Console.ReadKey();
				//            }
				//            //int keyResult = 0, valueResult = 0;
				//            //if(int.TryParse(valuelist[0], out keyResult))
				//            //{
				//            //    if(int.TryParse(valuelist[1], out valueResult))
				//            //    { 
				//            //    }
				//            //}
				//        }
				////    }
				//    value = dic;
				//}
				//#endregion

				if (value is DBNull)
				{
					value = getColumnDefault(sheet, column, firstDataRow);
				}
				else if (value is double)
				{ // 去掉数值字段的“.0”
					double num = (double)value;
					if ((int)num == num)
						value = (int)num;
				}

				string fieldName = column.ToString();
				// 表头自动转换成小写
				if (lowcase)
					fieldName = fieldName.ToLower();

				if (string.IsNullOrEmpty(fieldName))
					fieldName = string.Format("col_{0}", col);

				rowData[fieldName] = value;
				col++;
				index++;
			}
			return rowData;
		}

		/// <summary>
		/// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
		/// </summary>
		private static object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
		{
			for (int i = firstDataRow; i < sheet.Rows.Count; i++)
			{
				object value = sheet.Rows[i][column];
				Type valueType = value.GetType();
				if (valueType != typeof(System.DBNull))
				{
					if (valueType.IsValueType)
						return Activator.CreateInstance(valueType);
					break;
				}
			}
			return "";
		}
	}
}
