namespace ExcelReader.Tests;

using ExcelReader.Excel;
using ExcelReader.Sql;
using ExcelReader.Tests.Files;
using System.Data;

public class ImportMultiSheetTest
{
    const string ConnectionString = "Data Source=(LocalDb)\\MSSQLLocalDB;Initial Catalog=ExcelReader;Integrated Security=SSPI";

    [Fact]
    public async Task ImportTest()
    {
        foreach (var options in GetImportOptions())
        {
            var dataTable = ReadFile(options);

            await ImportFile(options, dataTable);
        }
    }
    
    private static DataTable ReadFile(ImportOptions options)
    {
        var reader = new ExcelSheetReader();

        using var fileStream = Resources.GetFileStream(options.ResourceName);
        var dataTable = reader.Read(fileStream, options);
        return dataTable;
    }

    private static async Task ImportFile(ImportOptions options, DataTable dataTable)
    {      
        var sqlWriter = new SqlDataWriter();

        await sqlWriter.ClearTableAsync(ConnectionString, options.TableName);
        
        await sqlWriter.WriteDataAsync(
            dataTable,
            new SqlDataWriterOptions
            {
                ConnectionString = ConnectionString,
                TableName = options.TableName
            });
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