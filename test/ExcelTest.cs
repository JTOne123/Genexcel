using Genexcel;
using Genexcel.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Genexcel.Tests {
	[TestClass]
	public class ExcelTest {
		public ExcelTest() {
			if (!Directory.Exists("C:/Tests/excel")) {
				Directory.CreateDirectory("C:/Tests/excel");
			}
		}
		//[TestMethod]
		//public void CreateZipWithFiles() {
		//	using (var builder = new ZipBuilder()) {
		//		builder.AddEntry("arquivo1.txt", "opa");
		//		builder.AddEntry("arquivo2.txt", "opa2");
		//		builder.Save("C:/Tests/test-create-zip-with-files.zip");
		//	}
		//}

		string CreateFileName(MethodBase mi) {
			return $"C:/Tests/excel/{mi.Name}-{DateTime.Now:yyyy-MM-dd-mm-ss}.xlsx";
		}

		void Save(Document excel, MethodBase mi) {
			excel.Save(CreateFileName(mi));
		}

		//[TestMethod]
		//public void CloneAndSaveZipTest() {
		//	var ass = typeof(Builder).GetTypeInfo().Assembly;
		//	using (var builder = new ZipBuilder(ass.GetManifestResourceStream($"{Builder.ResourcesPath}.ExcelTemplate.xlsx"))) {
		//		builder.Save("C:/Tests/test-clone-save-zip.xlsx");
		//	}
		//}

		[TestMethod]
		public void EmptyTest() {
			var excel = new Document();
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void EmptyTestStream() {
			var excel = new Document();
			excel.GetSheets().First().Name = "Teste";
			using (var memoryStream = excel.Save()) {
				var bytes = memoryStream.ToArray();
				File.WriteAllBytes(CreateFileName(MethodInfo.GetCurrentMethod()), bytes);
			}
			//Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void AreaChartTest() {
			var excel = new Document();
			var sheet = excel.GetSheets().First();
			sheet.Charts.Add(new AreaChart() {
				Data = new ChartData() {
					Labels = new List<string>() {
						DateTime.Now.ToShortDateString(),
						DateTime.Now.AddDays(1).ToShortDateString(),
						DateTime.Now.AddDays(2).ToShortDateString(),
					},
					Datasets = new List<ChartDataset>() {
						new ChartDataset() {
							Title = "Testando",
							Data = new List<decimal>() { 1, 2, 3 }
						}
					}
				}

			});
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void BarChartTest() {
			var excel = new Document();
			var sheet = excel.GetSheets().First();
			sheet.Charts.Add(new BarChart() {
				Data = new ChartData() {
					Labels = new List<string>() {
						"SP",
						"RJ",
						"BH"
					},
					Datasets = new List<ChartDataset>() {
						new ChartDataset() {
							Title = "Testando",
							Data = new List<decimal>() { 1, 2, 3 }
						}
					}
				}

			});
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void CreateSheetWithDataStreamTest() {
			var excel = new Document();
			excel.AddSheet("Test")
				.Add(new Cell(1, 1, "Test"));
			using (var memoryStream = excel.Save()) {
				var bytes = memoryStream.ToArray();
				File.WriteAllBytes(CreateFileName(MethodInfo.GetCurrentMethod()), bytes);
			}
		}


		[TestMethod]
		public void CreateSheetLargeNameTest() {
			var excel = new Document();
			excel.AddSheet("Teste Nova Sheet Com um nome bem grande");
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void CreateSheetWithDataTest() {
			var excel = new Document();
			var newSheet = excel.AddSheet("Test");
			newSheet.Add(new Cell(1, 1, "Test"));
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void RenameFirstSheetAndWriteData() {
			var excel = new Document();
			var sheet = excel.GetSheets().First();
			sheet.Name = "Minha planilha";
			sheet.Add(new Cell(1, 1, "Test"));
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void CreateSheetWithNumericTest() {
			var excel = new Document();
			excel.AddSheet("Test")
				.Add(new Cell(1, 1, 1));
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void ManySheetsManyCellsTest() {
			var excel = new Document();
			for (int i = 0; i < 20; i++) {
				var sheet = excel.AddSheet($"S{i}");
				for (int j = 1; j < 20; j++) {
					for (int k = 1; k < 20; k += 2) {
						sheet.Add(new Cell(j, k, $"Test{i}:{j}:{k}"));
					}
				}
			}
			Save(excel, MethodInfo.GetCurrentMethod());
		}


		[TestMethod]
		public void MuchRowsTest() {
			var excel = new Document();
			for (int i = 0; i < 1; i++) { //Sheets
				var sheet = excel.AddSheet($"S{i}");
				for (int j = 1; j < 20000; j++) { //Rows
					for (int k = 1; k < 20; k += 2) { //Columns
						sheet.Add(new Cell(j, k, $"Test{i}:{j}:{k}"));
					}
				}
			}
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void SetColumnWidthTest() {
			var excel = new Document();
			var sheet = excel.GetSheets().First();
			sheet.Add(new Cell(1, 1, "Test"));
			sheet.SetColumnWidth(100, 3, 5);
			// sheet.SetColumnWidth(50, 4, 10);
			//sheet.SetColumnWidth(200, 8, 20);
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void HyperlinkTest() {
			var excel = new Document();
			var sheet = excel.GetSheets().First();
			sheet.Add(new Cell(1, 1, "Test") { Hyperlink = "https://www.google.com" });
			Save(excel, MethodInfo.GetCurrentMethod());
		}

		[TestMethod]
		public void ReadMeExample1() {
			var excel = new Document();

			//Get the first sheet 
			var sheet1 = excel.GetSheets().First();
			//Change sheet name 
			sheet1.Name = "My first sheet";

			//Write some data 
			sheet1.Add(new Cell(1, 1, "Test 1"));

			//Create another sheet 
			var sheet2 = excel.AddSheet("My second sheet");

			//Write some data 
			sheet2.Add(new Cell(1, 1, "Test 2"));

			//Save to file, or stream... 
			excel.Save("myFile.xlsx");
		}
	}
}
