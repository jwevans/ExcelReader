namespace ExcelReader.Excel;

public class ExcelSheetReaderOptions
{
    public required string SheetName { get; init; }

    public List<ExcelColumn> Columns { get; init; } = new();

    public int SkipRows { get; init; } = 0;

    public List<string> NullValues { get; init; } = new();
}
