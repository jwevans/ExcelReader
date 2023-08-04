namespace ExcelReader.Sql;

using Microsoft.Data.SqlClient;
using System.Data;

public class SqlDataWriter : ISqlDataWriter
{
    private readonly Dictionary<string, Dictionary<string, DataColumn>> _tableCache = new(StringComparer.OrdinalIgnoreCase);

    public async Task ClearTableAsync(
        string connectionString,
        string tableName,
        CancellationToken cancellationToken = default)
    {
        using SqlConnection sqlConnection = await CreateConnection(connectionString, cancellationToken);

        var command = new SqlCommand($"TRUNCATE TABLE {tableName}", sqlConnection);

        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task WriteDataAsync(
        DataTable dataTable, 
        SqlDataWriterOptions options, 
        CancellationToken cancellationToken = default)
    {
        using SqlConnection sqlConnection = await CreateConnection(options.ConnectionString, cancellationToken);

        var tableColumns = GetTableColumns(sqlConnection, options.TableName);

        using var bulkCopy = new SqlBulkCopy(sqlConnection, SqlBulkCopyOptions.Default, null)
        {
            BatchSize = options.BatchSize,
            BulkCopyTimeout = (int)options.BulkCopyTimeout.TotalSeconds,
            DestinationTableName = options.TableName,
        };

        // add the additional columns to the source data table and populate the data rows
        foreach (var column in options.AdditionalColumns)
        {
            if (!dataTable.Columns.Contains(column.Name))
            {
                dataTable.Columns.Add(column.Name, column.DataType ?? typeof(string));

                foreach (DataRow dataRow in dataTable.Rows)
                {
                    dataRow[column.Name] = column.Value;
                }
            }
        }

        // add mappings from the source data to the database table names as bulk insert is case sensitive
        bulkCopy.ColumnMappings.Clear();
        foreach (DataColumn dataColumn in dataTable.Columns)
        {
            if (tableColumns.TryGetValue(dataColumn.ColumnName, out var tableColumn))
            {
                bulkCopy.ColumnMappings.Add(dataColumn.ColumnName, tableColumn.ColumnName);
            }
            else
            {
                Console.WriteLine($"Column {dataColumn.ColumnName} does not exist in database");                
            }
        }

        await bulkCopy.WriteToServerAsync(dataTable, cancellationToken);
    }

    private Dictionary<string, DataColumn> GetTableColumns(SqlConnection sqlConnection, string tableName)
    {
        if (_tableCache.TryGetValue(tableName, out var tableColumns))
        {
            return tableColumns;
        }

        using var sqlCommand = new SqlCommand
        {
            CommandText = $"SELECT * FROM {tableName} WHERE 0 = 1",
            CommandType = CommandType.Text,
            Connection = sqlConnection
        };

        var dataTable = new DataTable();
        using (var dataAdapter = new SqlDataAdapter(sqlCommand))
        {
            _ = dataAdapter.FillSchema(dataTable, SchemaType.Source);
        }

        tableColumns = dataTable.Columns
            .OfType<DataColumn>()
            .ToDictionary(
                k => k.ColumnName,
                v => v,
                StringComparer.OrdinalIgnoreCase);

        _tableCache.Add(tableName, tableColumns);

        return tableColumns;
    }

    private static async Task<SqlConnection> CreateConnection(string connectionString, CancellationToken cancellationToken)
    {
        var sqlConnection = new SqlConnection(connectionString);
        await sqlConnection.OpenAsync(cancellationToken);
        return sqlConnection;
    }
}
