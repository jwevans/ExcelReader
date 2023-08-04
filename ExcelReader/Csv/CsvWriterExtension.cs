namespace ExcelReader.Csv;

using CsvHelper;
using System.Data;

public static class CsvWriterExtension
{
    public static void WriteDataTable(this CsvWriter writer, DataTable dataTable, List<string>? excludeColumns = null)
    {
        // Get the columns to write out
        var csvColumns = dataTable.Columns.OfType<DataColumn>().ToList();
        if (excludeColumns is not null)
        {
            csvColumns.RemoveAll(c => excludeColumns.Contains(c.ColumnName, StringComparer.OrdinalIgnoreCase));
        }

        // Write the header
        foreach (var column in csvColumns)
        {
            writer.WriteField(column.ColumnName);
        }
        writer.NextRecord();

        // Write the rows
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (var column in csvColumns)
            {
                var value = row[column];

                writer.WriteField(value);
            }
            writer.NextRecord();
        }
    }
}