namespace ExcelReader.Sql;

using System.Collections.Generic;

public class SqlDataWriterOptions
{
    public required string TableName { get; init; }

    public required string ConnectionString { get; init; }

    public int BatchSize { get; init; } = 1000;

    public TimeSpan BulkCopyTimeout { get; init; } = TimeSpan.FromMinutes(3);

    public List<Column> AdditionalColumns { get; } = new();

    public record Column(string Name, object Value, Type? DataType = null);
}
