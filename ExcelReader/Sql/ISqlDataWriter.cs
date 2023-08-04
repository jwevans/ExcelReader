namespace ExcelReader.Sql;

using System.Data;

public interface ISqlDataWriter
{
    Task ClearTableAsync(
        string connectionName,
        string tableName,
        CancellationToken cancellationToken = default);

    Task WriteDataAsync(
        DataTable dataTable,
        SqlDataWriterOptions options,
        CancellationToken cancellationToken = default);
}
