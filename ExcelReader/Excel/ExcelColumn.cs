namespace ExcelReader.Excel;

public class ExcelColumn
{
    public ExcelColumn(int columnIndex, string mappingName)
    {
        ColumnIndex = columnIndex;
        MappingName = mappingName;
    }

    public ExcelColumn(string columnName, string? mappingName = null)
    {
        ColumnName = columnName;
        MappingName = mappingName ?? ColumnName;
    }

    public int? ColumnIndex { get; }
    public string? ColumnName { get; }
    public string MappingName { get; }
}