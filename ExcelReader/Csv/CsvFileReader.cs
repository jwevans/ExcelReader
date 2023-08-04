namespace ExcelReader.Csv;

using CsvHelper;
using CsvHelper.Configuration;
using System.Data;
using System.Globalization;
using Throw;

internal class CsvFileReader : ICsvFileReader
{
    public DataTable Read(string fullName, CsvFileReaderOptions options)
    {
        fullName.ThrowIfNull();
        options.ThrowIfNull();

        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            IgnoreBlankLines = true,
            TrimOptions = options.TrimValues ? TrimOptions.Trim : TrimOptions.None
        };

        using var streamReader = new StreamReader(fullName);
        using var csvReader = new CsvReader(streamReader, config);
        using var dataReader = new CsvDataReader(csvReader);

        csvReader.Context.TypeConverterOptionsCache.GetOptions<string>().NullValues.AddRange(options.NullValues);

        var dataTable = new DataTable();
        dataTable.Load(dataReader);

        return dataTable;
    }
}
