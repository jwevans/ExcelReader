namespace ExcelReader.Csv;

using System.Data;

public interface ICsvFileReader
{
    DataTable Read(string fullName, CsvFileReaderOptions options);
}
