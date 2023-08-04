using ExcelReader.Excel;
using ExcelReader.Sql;
using ExcelReader.Tests.Files;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Data;

namespace ExcelReader.Tests;

public class ImportSheetTest
{    
    [Fact]
    public async Task ImportTest()
    {
        var dataTable = ReadFile();

        await ImportFile(dataTable);
    }
    
    private static DataTable ReadFile()
    {
        var reader = new ExcelSheetReader();
        var options = GetExcelReaderOptions();
        var fileStream = Resources.GetFileStream(Resources.FinancialSample);

        var dataTable = reader.Read(fileStream, options);

        return dataTable;
    }

    private static async Task ImportFile(DataTable dataTable)
    {
        const string ConnectionString = "Data Source=(LocalDb)\\MSSQLLocalDB;Initial Catalog=ExcelReader;Integrated Security=SSPI";
        const string TableName = "Financial";

        var sqlWriter = new SqlDataWriter();

        await sqlWriter.ClearTableAsync(ConnectionString, TableName);
        
        await sqlWriter.WriteDataAsync(
            dataTable,
            new SqlDataWriterOptions
            {
                ConnectionString = ConnectionString,
                TableName = TableName
            });
    }

    private static ExcelSheetReaderOptions GetExcelReaderOptions()
    {
        return new()
        {
            SheetName = "Sheet1",                
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
    }
}