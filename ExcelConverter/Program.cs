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
            //IL_002e: Unknown result type (might be due to invalid IL or missing references)
            //IL_0034: Expected O, but got Unknown
            try
            {
                using (FileStream fileStream = File.Open(options.InputFile, FileMode.Open, FileAccess.Read))
                {
                    using (StreamWriter streamWriter = new StreamWriter(File.Open(options.TargetFile, FileMode.Create), Encoding.Unicode))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            using (JsonTextWriter jsonWriter = new JsonTextWriter(streamWriter))
                            {
                                while (reader.Name.ToLower() != options.SheetName.ToLower())
                                {
                                    reader.NextResult();
                                }
                                jsonWriter.Formatting = Formatting.Indented;
                                for (int i = 0; i < options.SkipRows; i++)
                                {
                                    reader.Read();
                                }
                                List<string> titles = GetOrCreateHeader(options, reader);
                                ((JsonWriter)jsonWriter).WriteStartArray();
                                while (reader.Read())
                                {
                                    if (options.MainColumn.HasValue)
                                    {
                                        if (string.IsNullOrEmpty(reader.GetString(options.MainColumn.Value)))
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

                                    for (int column = 0; column < columnsCount; column++)
                                    {
                                        jsonWriter.WritePropertyName(titles[column]);
                                        jsonWriter.WriteValue((reader)[column] ?? string.Empty);
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
                Console.WriteLine("Program filed with message:" + ex.Message);
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
                    string title = reader.GetString(i);
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
                if (!string.IsNullOrWhiteSpace(reader.GetString(i)))
                {
                    result = false;
                    break;
                }
            }
            return result;
        }
    }
}