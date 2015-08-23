using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Text;

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

			// 读取参数，判断是否应直接退出
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
					diOutput = new DirectoryInfo(strOutputDir);
				}
				catch (Exception ex)
				{
					Debug.WriteLine(ex.Message);
					Console.WriteLine(ex.Message);
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

						Console.WriteLine("====== {0} ======", strWorksheetName);

						// 确保输出路径存在
						string strWorksheetOutputDir = strOutputDir + "\\" + strWorksheetName;
						DirectoryInfo diWorksheetOutput = new DirectoryInfo(strWorksheetOutputDir);

						while (!diWorksheetOutput.Exists)
						{
							try
							{
								diWorksheetOutput.Create();
								Thread.Sleep(200);
								diWorksheetOutput = new DirectoryInfo(strWorksheetOutputDir);
							}
							catch (Exception ex)
							{
								Debug.WriteLine(ex.Message);
								Console.WriteLine(ex.Message);
								Console.WriteLine("输出目录创建失败。");

								// 退出程序
								if (bReadKeyBeforeExit)
								{
									Console.ReadKey();
									return;
								}
							}
						}

						List<string> listRowNames = new List<string>();
						List<string> listColumnNames = new List<string>();
						int nRowNo = 1;
						int nColNo = 2;

						listColumnNames.Add("Column_0");
						listColumnNames.Add("Column_1");

						while (true)
						{
							object objCellData = worksheet.Cells[nRowNo, nColNo++].Value;

							if (null == objCellData)
							{
								break;
							}

							string strCellData = objCellData.ToString().Trim();

							if (0 == strCellData.Length)
							{
								break;
							}

							listColumnNames.Add(strCellData);
						}

						nRowNo = 2;
						nColNo = 1;
						listRowNames.Add("Row_0");
						listRowNames.Add("Row_1");

						while (true)
						{
							object objCellData = worksheet.Cells[nRowNo++, nColNo].Value;

							if (null == objCellData)
							{
								break;
							}

							string strCellData = objCellData.ToString().Trim();

							if (0 == strCellData.Length)
							{
								break;
							}

							listRowNames.Add(strCellData);
						}

						string[] strRowNames = listRowNames.ToArray();
						string[] strColumnNames = listColumnNames.ToArray();

						listRowNames.Clear();
						listColumnNames.Clear();
						listRowNames = null;
						listColumnNames = null;

						for (nColNo = 2; nColNo < strColumnNames.Length; nColNo++)
						{
							Dictionary<string, object> dicJsonRoot = new Dictionary<string, object>();

							for (nRowNo = 2; nRowNo < strRowNames.Length; nRowNo++)
							{
								object objCellData = worksheet.Cells[nRowNo, nColNo].Value;

								if (null == objCellData)
								{
									dicJsonRoot.Add(strRowNames[nRowNo], objCellData);
								}
								else
								{
									Type typeCellData = objCellData.GetType();

									if (typeof(float) == typeCellData || typeof(double) == typeCellData)
									{
										Double nCellData = Convert.ToDouble(objCellData);
										nCellData -= Math.Truncate(nCellData);

										if (Math.Abs(nCellData) < 0.0001)
										{
											dicJsonRoot.Add(strRowNames[nRowNo], Convert.ToInt64(objCellData));
										}
										else
										{
											dicJsonRoot.Add(strRowNames[nRowNo], objCellData);
										}
									}
									else
									{
										dicJsonRoot.Add(strRowNames[nRowNo], objCellData);
									}
								}
							}

							Console.WriteLine("--- {0}.json ---", strColumnNames[nColNo]);
							Console.WriteLine(JsonConvert.SerializeObject(dicJsonRoot, Formatting.Indented));

							using (FileStream fsJson = File.Create(strWorksheetOutputDir + "\\" + strColumnNames[nColNo] + ".json"))
							{
								byte[] byteJsonData = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(dicJsonRoot, Formatting.None));

								try
								{
									fsJson.Write(byteJsonData, 0, byteJsonData.Length);
								}
								catch (Exception ex)
								{
									Debug.WriteLine(ex.Message);
									Console.WriteLine(ex.Message);
								}
							}
						}
					}
				}
			}

			Console.WriteLine("转换完成。");
			// 退出程序
			if (bReadKeyBeforeExit)
			{
				Console.ReadKey();
			}
		}
	}
}
