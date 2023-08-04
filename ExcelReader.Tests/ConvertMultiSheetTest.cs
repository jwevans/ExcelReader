namespace ExcelReader.Tests;

using CsvHelper;
using ExcelReader.Csv;
using ExcelReader.Excel;
using ExcelReader.Tests.Files;
using System.Data;
using System.Globalization;

public class ConvertMultiSheetTest
{
    [Fact]
    public void ConvertTest()
    {
        foreach (var options in GetImportOptions())
        {
            var dataTable = ReadFile(options);

            WriteFile(options, dataTable);
        }
    }
    
    private static DataTable ReadFile(ImportOptions options)
    {
        var reader = new ExcelSheetReader();

        using var fileStream = Resources.GetFileStream(options.ResourceName);
        var dataTable = reader.Read(fileStream, options);
        return dataTable;
    }

    private static void WriteFile(ImportOptions options, DataTable dataTable)
    {
        var csvFilename = Path.Combine($@"C:\aaTemp\{options.SheetName}.csv");

        using var writer = new StreamWriter(csvFilename) { AutoFlush = true };
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        csv.Context.TypeConverterOptionsCache.GetOptions<DateTime>().Formats = new[] { "dd-MM-yyyy" };

        csv.WriteDataTable(dataTable);
    }  

    private static IEnumerable<ImportOptions> GetImportOptions()
    {
        yield return new ImportOptions
        {
            TableName = "Financial",
            ResourceName = Resources.MultiSheetSample,
            SheetName = "Financials",
            NullValues = new() { "N.A." },
            Columns =
            {
                new ExcelColumn("Segment"),
                new ExcelColumn("Country"),
                new ExcelColumn("Units Sold"),
                new ExcelColumn("Manufacturing Price"),
                new ExcelColumn("Sales"),
                new ExcelColumn("Profit"),
                new ExcelColumn("Date"),
                new ExcelColumn("Month Number"),
            }
        };

        yield return new ImportOptions
        {
            TableName = "Employee",
            ResourceName = Resources.MultiSheetSample,
            SheetName = "Employees",
            NullValues = new() { "N.A." },
            Columns =
            {
                new ExcelColumn("EEID"),
                new ExcelColumn("Full Name"),
                new ExcelColumn("Job Title"),
                new ExcelColumn("Hire Date"),
                new ExcelColumn("Department"),
            }
        };
    }

    private class ImportOptions : ExcelSheetReaderOptions
    {
        public required string TableName { get; init; }
        public required string ResourceName { get; init; }
    }
}