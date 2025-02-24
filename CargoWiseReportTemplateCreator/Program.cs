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

			if (opts.WhereClause != "")
				opts.WhereClause = $"WHERE {opts.WhereClause}";

			if (!File.Exists(opts.OutputFile))
				CreateNewBlankExcelFile(opts);
			else
				AppendToExistingExcelFile(opts);
		}

		private static void CreateNewBlankExcelFile(Options opts)
		{
			MemoryStream ms = new MemoryStream();
			using (ExcelPackage excelPackage = new ExcelPackage())
			{
				CreateWorksheet(ms, excelPackage, opts);
			}

			ms.Position = 0;
			using (FileStream file = new FileStream(opts.OutputFile, FileMode.Create, FileAccess.Write))
				ms.WriteTo(file);
		}

		private static void AppendToExistingExcelFile(Options opts)
		{
			MemoryStream ms = new MemoryStream();
			using (FileStream fs = File.OpenRead(opts.OutputFile))
			{
				using (ExcelPackage excelPackage = new ExcelPackage(fs))
				{
					CreateWorksheet(ms, excelPackage, opts);
				}
			}

			ms.Position = 0;
			using (FileStream file = new FileStream(opts.OutputFile, FileMode.Create, FileAccess.Write))
				ms.WriteTo(file);
		}

		private static void CreateWorksheet(MemoryStream ms, ExcelPackage from, Options opts)
		{
			ExcelWorkbook excelWorkBook = from.Workbook;
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
			from.SaveAs(ms);
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