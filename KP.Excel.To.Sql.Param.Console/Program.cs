using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using KP.Excel.To.Sql.Param.Console.Models;

namespace KP.Excel.To.Sql.Param.Console
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var records = ReadWorksheetRecords();
            System.Console.WriteLine($"Attempting to parse {records.Count()} records.");

            foreach (var record in records)
            {
                System.Console.WriteLine($"\'{record.Account}\',");
            }

            System.Console.WriteLine("Press any key to close");
            System.Console.ReadLine();
        }

        private static IEnumerable<Worksheet> ReadWorksheetRecords()
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                PrepareHeaderForMatch = args => args.Header.ToLower(),
            };

            using var reader = new StreamReader("..\\..\\..\\..\\worksheet.csv");
            using var csv = new CsvReader(reader, config);
            return csv.GetRecords<Worksheet>().ToList();
        }
    }
}