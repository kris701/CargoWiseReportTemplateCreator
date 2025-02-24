using CommandLine;

namespace CargoWiseReportTemplateCreator
{
	public class Options
	{
		[Option('t', "table", Required = true, HelpText = "What table you want a report from.")]
		public string TableName { get; set; }
		[Option('c', "columns", Required = true, HelpText = "What table you want a report from.")]
		public IEnumerable<string> ColumnNames { get; set; } = new List<string>();

		[Option('w', "where", Required = false, HelpText = "The additional where clause to add to the select statement.")]
		public string WhereClause { get; set; } = "";
		[Option('o', "output", Required = false, HelpText = "Name of the output file", Default = "out.xlsx")]
		public string OutputFile { get; set; } = "out.xlsx";
		[Option('d', "delete", Required = false, HelpText = "If the program should remove an existing output file or append to it instead.", Default = false)]
		public bool DeleteExisting { get; set; } = false;
	}
}
