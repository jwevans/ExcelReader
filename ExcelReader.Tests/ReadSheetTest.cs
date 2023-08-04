using ExcelReader.Excel;
using ExcelReader.Tests.Files;

namespace ExcelReader.Tests;

public class ReadSheetTest
{
    [Fact]
    public void ReadTest()
    {
        var reader = new ExcelSheetReader();
        var options = GetExcelReaderOptions();

        using var fileStream = Resources.GetFileStream(Resources.FinancialSample);
        
        var dataTable = reader.Read(fileStream, options);

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