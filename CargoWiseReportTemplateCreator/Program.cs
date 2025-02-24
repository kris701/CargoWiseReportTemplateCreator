using CommandLine;
using CommandLine.Text;
using OfficeOpenXml;

namespace CargoWiseReportTemplateCreator
{
	internal class Program
	{
		public static async Task Main(string[] args)
		{
			var parser = new Parser(with => with.HelpWriter = null);
			var parserResult = parser.ParseArguments<Options>(args);
			parserResult.WithNotParsed(errs => DisplayHelp(parserResult, errs));
			await parserResult.WithParsedAsync(Run);
		}

		public static async Task Run(Options opts)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			if (opts.DeleteExisting && File.Exists(opts.OutputFile))
				File.Delete(opts.OutputFile);

			if (!File.Exists(opts.OutputFile))
				CreateNewBlankExcelFile(opts.OutputFile);

			if (opts.WhereClause != "")
				opts.WhereClause = $"WHERE {opts.WhereClause}";

			MemoryStream ms = new MemoryStream();
			using (FileStream fs = File.OpenRead(opts.OutputFile))
			{
				using (ExcelPackage excelPackage = new ExcelPackage(fs))
				{
					ExcelWorkbook excelWorkBook = excelPackage.Workbook;
					var sheet = excelWorkBook.Worksheets.Add(opts.TableName);
					sheet.SetValue(1, 1, "#config");
					sheet.SetValue(2, 1, "PageStyle=Continuous");
					sheet.SetValue(3, 1, $"Data:ReportData=SELECT * FROM {opts.TableName} {opts.WhereClause}");
					sheet.SetValue(4, 1, "#DocumentHeader");
					var offset = 2;
					foreach (var col in opts.ColumnNames)
						sheet.SetValue(5, offset++, col);
					sheet.SetValue(6, 1, "#SectionBody:Data=ReportData");
					offset = 2;
					foreach (var col in opts.ColumnNames)
						sheet.SetValue(7, offset++, $"<ReportData.{col}>");
					sheet.SetValue(8, 1, "#endofreport");
					excelPackage.SaveAs(ms);
				}
			}

			ms.Position = 0;
			using (FileStream file = new FileStream(opts.OutputFile, FileMode.Create, FileAccess.Write))
				ms.WriteTo(file);
		}

		private static void CreateNewBlankExcelFile(string name)
		{
			MemoryStream ms = new MemoryStream();
			using (ExcelPackage excelPackage = new ExcelPackage())
			{
				ExcelWorkbook excelWorkBook = excelPackage.Workbook;
				var sheet = excelWorkBook.Worksheets.Add("Filter");
				sheet.SetValue(1, 1, "Batch");
				sheet.SetValue(1, 2, "Type");
				sheet.SetValue(1, 3, "Number");
				sheet.SetValue(2, 1, "#End");
				excelPackage.SaveAs(ms);
			}

			ms.Position = 0;
			using (FileStream file = new FileStream(name, FileMode.Create, FileAccess.Write))
				ms.WriteTo(file);
		}

		private static void HandleParseError(IEnumerable<CommandLine.Error> errs)
		{
			var sentenceBuilder = SentenceBuilder.Create();
			foreach (var error in errs)
				if (error is not HelpRequestedError)
					Console.WriteLine(sentenceBuilder.FormatError(error));
		}

		private static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<CommandLine.Error> errs)
		{
			var helpText = HelpText.AutoBuild(result, h =>
			{
				h.AddEnumValuesToHelpText = true;
				return h;
			}, e => e, verbsIndex: true);
			Console.WriteLine(helpText);
			HandleParseError(errs);
		}
	}
}