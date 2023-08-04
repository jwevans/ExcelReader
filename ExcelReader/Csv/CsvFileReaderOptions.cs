namespace ExcelReader.Csv;

public class CsvFileReaderOptions
{
    public bool TrimValues { get; set; } = true;
    public List<string> NullValues { get; init; } = new() { string.Empty };
}
