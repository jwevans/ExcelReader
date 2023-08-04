using ExcelReader.Excel;
using ExcelReader.Tests.Files;

namespace ExcelReader.Tests;

public class ReadSheetTest
{
    [Fact]
    public void Test1()
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
            }
        };
    }
}