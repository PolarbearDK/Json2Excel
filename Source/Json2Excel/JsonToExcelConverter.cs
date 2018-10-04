using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Miracle.Arguments;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace Json2Excel
{
	[ArgumentSettings(StartOfArgument = new []{'-'})]
	[ArgumentDescription("Convert Json file to excel file.")]
	class JsonToExcelConverter
	{
		[ArgumentName("Input", "I")]
		[ArgumentRequired]
		[ArgumentDescription("Input json file.")]
		public string Input { get; set; }

		[ArgumentName("Output", "O")]
		[ArgumentDescription("Output excel file. Default: Import file with extension specified in Extension argument")]
		public string OutputFileName { get; set; }

		[ArgumentName("Extension", "E")]
		[ArgumentDescription("Output excel file. Default: Import file with .xlsx extension")]
		public string OutputExtension { get; set; } = ".xlsx";

		[ArgumentName("Sample", "S")]
		[ArgumentDescription("Number of rows to sample for json properties. Default: all rows")]
		public int Sample { get; set; } = int.MaxValue;

		[ArgumentName("Worksheet", "W")]
		[ArgumentDescription("Name of worksheet. Default is the filename of the input file.")]
		public string Worksheet { get; set; }

		[ArgumentName("Help", "H", "?")]
		[ArgumentHelp()]
		[ArgumentDescription("Show help.")]
		public bool Help { get; set; }

		public void Convert()
		{
			// read JSON directly from a file
			using (StreamReader file = File.OpenText(Input))
			{
				using (JsonTextReader reader = new JsonTextReader(file))
				{
					JArray array = (JArray)JToken.ReadFrom(reader);

					var properties = new PropertyPath[] { };
					foreach (JObject token in array.Take(Sample))
					{
						properties = properties.Union(PropertyPath.GetPropertyPaths(token)).ToArray();
					}

					using (var p = new ExcelPackage())
					{
						var ws = p.Workbook.Worksheets.Add(GetWorksheetName());

						// Column headers
						int col = 1;
						int row = 1;
						foreach (var property in properties)
						{
							ws.SetValue(row, col++, property.ToString());
						}

						// Data rows
						foreach (JObject token in array)
						{
							row++;
							col = 1;
							foreach (var property in properties)
							{
								if (property.GetValue(token) is JValue jValue)
								{
									var value = jValue.Value;
									ws.SetValue(row, col, value);
									if (value is DateTime)
									{
										ws.Cells[row,col].Style.Numberformat.Format = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
									}
								}

								col++;
							}
						}

						// Auto fit all columns
						ws.Cells[1, 1, row, col].AutoFitColumns();

						// And save result
						p.SaveAs(new FileInfo(GetOutputFileName()));
					}
				}
			}
		}

		private string GetWorksheetName()
		{
			return string.IsNullOrWhiteSpace(Worksheet)
				? Path.GetFileName(Input)
				: Worksheet;
		}

		private string GetOutputFileName()
		{
			return string.IsNullOrWhiteSpace(OutputFileName)
				? Path.ChangeExtension(Input, OutputExtension)
				: OutputFileName;
		}
	}
}