using CommandLine;

namespace ExcelConverter
{
    partial class Program
    {
        class Options
        {
            [Option('s', "source", Required = true, HelpText = "Input file to be converted.")]
            public string InputFile { get; set; }
            [Option('t', "target", Required = true, HelpText = "Target file to output json result.")]
            public string TargetFile { get; set; }

            [Option('c', "columns", Required = true, HelpText = "How many columns to read.")]
            public int Columns { get; set; }

            [Option("sheet", Required = true, HelpText = "Name of sheet to read.")]
            public string SheetName { get; set; }

            [Option("skipRows", Required = false, Default = 0, HelpText = "How many rows to skip.")]
            public int SkipRows { get; set; }

            [Option('h', "withHeader",  HelpText = "First reading row is a header. If false - set column leteral as header")]
            public bool WithHeader { get; set; }

            [Option('i', "withIdColumn", HelpText = "Add column for identity rows")]
            public bool WithIdColumn { get; set; }
            [Option("IdentityRowColumnName", Default = "ExcelRowId", HelpText = "Column name for identity row, default = \"ExcelRowId\"")]
            public string IdentityRowColumnName { get; set; }

            [Option('m', "mainColumn", Required = false, HelpText = "Column where check for existing row, if not specified check first columns(setted by CheckColumns). If this column are empty - close reader and save result")]
            public int? MainColumn { get; set; }

            [Option("checkColumns", Required = false, Default = 5, HelpText = "count of first N columns to check for existing row, default first 5 columns. If this columns are empty - close reader and save result")]
            public int CheckColumns { get; set; }
        }
    }
}