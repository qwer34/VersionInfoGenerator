using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VersionInfoGenerator
{
	class Program
	{
		static void Main(string[] args)
		{
			// 打印信息
			Console.WriteLine("************************************");
			Console.WriteLine("* {0} Ver. {1}", Application.ProductName, Application.ProductVersion);
			Console.WriteLine("* Powered by Xin Zhang");
			Console.WriteLine("* {0}", System.IO.File.GetLastWriteTime(Application.ExecutablePath));
			Console.WriteLine("* 启动路径下所有的Excel文件将被转换为json。");
			Console.WriteLine("*");
			Console.WriteLine("* 用法:");
			Console.WriteLine("* -----");
			Console.WriteLine("* {0} [-x]", Application.ProductName);
			Console.WriteLine("*  -x 转换完毕不等待用户按键，直接退出。");
			Console.WriteLine("************************************");

			// 读取参数，判断是否莹直接退出
			bool bReadKeyBeforeExit = true;

			foreach (string arg in args)
			{
				if ("-x" == arg.ToLower())
				{
					bReadKeyBeforeExit = false;
					break;
				}
			}

			// 确保输入路径存在
			DirectoryInfo diInput = new DirectoryInfo(Application.StartupPath);

			if (!diInput.Exists)
			{
				Console.WriteLine("启动路径不正确。");

				// 退出程序
				if (bReadKeyBeforeExit)
				{
					Console.ReadKey();
					return;
				}
			}

			// 确保输出路径存在
			string strOutputDir = diInput.FullName + "\\versioninfo";
			DirectoryInfo diOutput = new DirectoryInfo(strOutputDir);

			while (!diOutput.Exists)
			{
				try
				{
					diOutput.Create();
					Thread.Sleep(200);
				}
				catch (Exception Ex)
				{
					Debug.WriteLine(Ex.Message);
					Console.WriteLine(Ex.Message);
					Console.WriteLine("输出目录创建失败。");

					// 退出程序
					if (bReadKeyBeforeExit)
					{
						Console.ReadKey();
						return;
					}
				}
			}

			// 遍历、转换
			FileInfo[] fis = diInput.GetFiles();

			foreach (FileInfo fi in fis)
			{
				if (!fi.Exists)
				{
					continue;
				}

				if (".xlsx" != fi.Extension && ".xls" != fi.Extension)
				{
					continue;
				}

				using (ExcelPackage package = new ExcelPackage(fi))
				{
					foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
					{
						string strWorksheetName = worksheet.Name;
						string strWorksheetLowerName = strWorksheetName.ToLower();

						if (strWorksheetLowerName.StartsWith("模板") || strWorksheetLowerName.StartsWith("template"))
						{
							continue;
						}

						Console.WriteLine("====== {0} ======", strWorksheetLowerName);

						List<string> RowNameList = new List<string>();
						List<string> ColumnNameList = new List<string>();
						int RowNo = 1;
						int ColNo = 2;

						ColumnNameList.Add("Column_0");
						ColumnNameList.Add("Column_1");

						while (true)
						{
							object CellData = worksheet.Cells[RowNo, ColNo++].Value;

							if (null == CellData)
							{
								break;
							}

							string CellDataString = CellData.ToString().Trim();

							if (0 == CellDataString.Length)
							{
								break;
							}

							ColumnNameList.Add(CellDataString);
						}

						RowNo = 2;
						ColNo = 1;
						RowNameList.Add("Row_0");
						RowNameList.Add("Row_1");

						while (true)
						{
							object CellData = worksheet.Cells[RowNo++, ColNo].Value;

							if (null == CellData)
							{
								break;
							}

							string CellDataString = CellData.ToString().Trim();

							if (0 == CellDataString.Length)
							{
								break;
							}

							RowNameList.Add(CellDataString);
						}

						string[] RowNames = RowNameList.ToArray();
						string[] ColumnNames = ColumnNameList.ToArray();

						RowNameList.Clear();
						ColumnNameList.Clear();
						RowNameList = null;
						ColumnNameList = null;

						for (ColNo = 2; ColNo < ColumnNames.Length; ColNo++)
						{
							Dictionary<string, object> dicJsonRoot = new Dictionary<string, object>();

							for (RowNo = 2; RowNo < RowNames.Length; RowNo++)
							{
								dicJsonRoot.Add(RowNames[RowNo], worksheet.Cells[RowNo, ColNo].Value);
							}

							Console.WriteLine("--- {0}.json ---", ColumnNames[ColNo]);
							Console.WriteLine(JsonConvert.SerializeObject(dicJsonRoot, Formatting.Indented));
						}
					}
				}
			}

			// 退出程序
			if (bReadKeyBeforeExit)
			{
				Console.ReadKey();
			}
		}
	}
}
