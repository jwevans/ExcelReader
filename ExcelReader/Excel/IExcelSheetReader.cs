namespace ExcelReader.Excel;

using System.Data;

public interface IExcelSheetReader
{    
    DataTable Read(string filename, ExcelSheetReaderOptions options);

    DataTable Read(Stream fileStream, ExcelSheetReaderOptions options);
}
