using CommandLine;
using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter
{
    partial class Program
    {

        private static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithNotParsed(HandleParseError)
                .WithParsed(RunOptions);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private static void HandleParseError(IEnumerable<Error> obj)
        {
            Console.WriteLine("Please provide correct args!");
            Environment.Exit(-1);
        }

        private static void RunOptions(Options options)
        { 
            int columnNumber = -1;
            int rowNumber = -1;
            try
            {
                using (FileStream fileStream = File.Open(options.InputFile, FileMode.Open, FileAccess.Read))
                {
                    using (StreamWriter streamWriter = new StreamWriter(File.Open(options.TargetFile, FileMode.Create), Encoding.Unicode))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            using (JsonTextWriter jsonWriter = new JsonTextWriter(streamWriter) { Formatting = Formatting.Indented })
                            {

                                while (reader.Name.ToLower() != options.SheetName.ToLower())
                                {
                                    reader.NextResult();
                                }
                                for (rowNumber = 0; rowNumber < options.SkipRows; rowNumber++)
                                {
                                    reader.Read();
                                }
                                List<string> titles = GetOrCreateHeader(options, reader);
                                jsonWriter.WriteStartArray();
                                while (reader.Read())
                                {
                                    rowNumber++;
                                    if (options.MainColumn.HasValue)
                                    {
                                        if (string.IsNullOrEmpty(reader[(options.MainColumn.Value)]?.ToString()))
                                        {
                                            break;
                                        }
                                    }
                                    else if (IsEmptyColumns(options, reader))
                                    {
                                        break;
                                    }
                                    jsonWriter.WriteStartObject();

                                    var columnsCount = reader.FieldCount < options.Columns ? reader.FieldCount : options.Columns;
                                    for (  columnNumber = 0; columnNumber < columnsCount; columnNumber++)
                                    {
                                        jsonWriter.WritePropertyName(titles[columnNumber]);
                                        jsonWriter.WriteValue((reader)[columnNumber] ?? string.Empty);
                                    }
                                    if (options.WithIdColumn)
                                    {
                                        jsonWriter.WritePropertyName(options.IdentityRowColumnName);
                                        jsonWriter.WriteValue(rowNumber); 
                                    }
                                    jsonWriter.WriteEndObject();
                                }
                                jsonWriter.WriteEndArray();
                            }
                        }
                    }
                }
                Console.WriteLine($"Sheet {options.SheetName} from file {options.InputFile} parsed succesfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Program failed on row {rowNumber} and column {columnNumber} with message:{ex.Message}");
                Environment.Exit(-1);
            }
        }

        private static List<string> GetOrCreateHeader(Options options, IExcelDataReader reader)
        {
            List<string> list = new List<string>();
   
            if (options.WithHeader)
            {
                reader.Read();
                var columnsCount = reader.FieldCount < options.Columns ? reader.FieldCount : options.Columns;
                for (int i = 0; i < columnsCount; i++)
                {
                    string title = reader[i]?.ToString();
                    title = string.IsNullOrWhiteSpace(title) ? ((char)(65 + i)).ToString() : title;
                    list.Add(list.Any((string x) => x.Equals(title)) ? (title + $"({(char)(65 + i)})") : title);
                }
            }
            else
            {
                for (int j = 0; j < options.Columns; j++)
                {
                    list.Add($"{(char)(65 + j)}");
                }
            }
            return list;
        }

        private static bool IsEmptyColumns(Options options, IExcelDataReader reader)
        {
            bool result = true;
            for (int i = 0; i < options.CheckColumns; i++)
            {
                if (!string.IsNullOrWhiteSpace(reader[i]?.ToString()))
                {
                    result = false;
                    break;
                }
            }
            return result;
        }
    }
}