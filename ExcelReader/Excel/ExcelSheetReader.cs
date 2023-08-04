namespace ExcelReader.Excel;

using OfficeOpenXml;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

public class ExcelSheetReader : IExcelSheetReader
{
    public DataTable Read(Stream fileStream, ExcelSheetReaderOptions options)
    {
        using var excel = new ExcelPackage(fileStream);
        using var sheet = GetSheet(options.SheetName, excel);

        var headerRow = FindHeaderRow(sheet, options);
        var headers = GetHeaderColumns(sheet, options, headerRow);

        var dataTable = new DataTable();

        foreach (var header in headers)
        {
            dataTable.Columns.Add(header.MappingName);
        }

        var startRow = headerRow + 1 + options.SkipRows;
        for (var row = startRow; row <= sheet.Dimension.End.Row; row++)
        {
            var dataRow = dataTable.NewRow();
            foreach (var header in headers)
            {
                var value = GetCellValue(sheet, row, header.ColumnIndex, options.NullValues);

                dataRow.SetField(header.MappingName, value);
            }
            dataTable.Rows.Add(dataRow);
        }

        return dataTable;
    }

    public DataTable Read(string filename, ExcelSheetReaderOptions options)
    {
        using var fileStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        return Read(fileStream, options);
    }

    private static int FindHeaderRow(ExcelWorksheet sheet, ExcelSheetReaderOptions options)
    {
        for (var row = 1; row <= sheet.Dimension.End.Row; row++)
        {
            for (var col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                var value = GetCellValue(sheet, row, col);

                if (value is not null && options.Columns.Exists(c => ColumnNameComparer.Compare(c.ColumnName, value)))
                {
                    return row;
                }
            }
        }

        throw new ApplicationException("Could not find the header row");
    }

    private static List<HeaderColumn> GetHeaderColumns(ExcelWorksheet sheet, ExcelSheetReaderOptions options, int headerRow)
    {
        var columns = new List<HeaderColumn>();

        foreach (var column in options.Columns)
        {
            if (column.ColumnIndex.HasValue)
            {
                // if we have a column which has an index just use the index
                columns.Add(
                    new HeaderColumn(
                        column.ColumnIndex.Value,
                        column.MappingName,
                        column.MappingName));
            }
            else
            {
                // otherwise we will need to lookup the index from the column name
                int? columnIndex = null;

                for (var col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    var value = GetCellValue(sheet, headerRow, col);

                    if (!string.IsNullOrEmpty(value) && ColumnNameComparer.Compare(column.ColumnName, value))
                    {
                        columnIndex = col;
                        break;
                    }
                }

                if (columnIndex is null)
                {
                    throw new KeyNotFoundException(
                        $"The expected column '{column}' was not found in worksheet {sheet.Name} (searched on row index {headerRow}).");
                }

                columns.Add(
                    new HeaderColumn(
                        columnIndex.Value,
                        column.ColumnName!,
                        column.MappingName));
            }
        }

        return columns;
    }

    private static ExcelWorksheet GetSheet(string sheetName, ExcelPackage package)
    {
        var sheet = package
            .Workbook
            .Worksheets
            .Where(sheet => sheet.Name.Contains(sheetName, StringComparison.InvariantCultureIgnoreCase))
            .FirstOrDefault();

        if (sheet is null)
        {
            throw new ArgumentException($"Could not find sheet {sheetName} in excel file ");
        }

        return sheet;
    }

    private static string? GetCellValue(ExcelWorksheet worksheet, int row, int col, List<string>? nullValues = null)
    {
        var cell = worksheet.Cells[row, col];

        if (cell.Value == null)
        {
            return null;
        }

        if (cell.TryParseDate(out var date))
        {
            return date.ToCsvDateFormat();
        }

        var stringValue = cell.Value.ToString();

        if (nullValues is not null && stringValue is not null && nullValues.Contains(stringValue))
        {
            return null;
        }

        return stringValue;
    }

    public class ColumnNameComparer : IEqualityComparer<string?>
    {
        private static readonly ColumnNameComparer Comparer = new ColumnNameComparer();

        public static bool Compare(string? first, string? second)
        {
            return Comparer.Equals(first, second);
        }

        public bool Equals(string? first, string? second)
        {
            if (first is null && second is null)
            {
                return true;
            }

            if (first is null || second is null)
            {
                return false;
            }

            first = RemoveSpaces(first);
            second = RemoveSpaces(second);

            return StringComparer.OrdinalIgnoreCase.Equals(first, second);
        }

        public int GetHashCode(string obj)
        {
            return obj.GetHashCode();
        }

        private static string RemoveSpaces(string value)
        {
            return Regex.Replace(value, @"\s+", string.Empty);
        }
    }

    private record HeaderColumn(int ColumnIndex, string ColumnName, string MappingName);
}
