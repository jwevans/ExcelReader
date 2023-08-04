namespace ExcelReader.Excel;

using OfficeOpenXml;
using System.Text.RegularExpressions;

internal static class ExcelExtensions
{
    public const string CsvDateFormat = "dd/MM/yyyy hh:mm:ss tt";

    public static bool TryParseDate(this ExcelRange cell, out DateTime date)
    {
        //dd/mm/yyyy
        const string pattern = "^(?i)d{1,2}\\/m{1,2}\\/y{2,4}";

        // If one of the built in date formats is used for this cell OfficeOpenXml will try and
        // convert the value to be a DateTime data type in which case we know we have a date
        if (cell.Value is DateTime dateValue)
        {
            date = dateValue;
            return true;
        }

        // If a custom date format is used for this cell OfficeOpenXml isn't able to convert the
        // data type to be a DateTime so we need to try and detect if the value is a date based
        // on the format string. We are expecting "m/d/yyyy" in the excel file so use regex to
        // cope with the following variations : m/d/yyyy, mm/dd/yyyy, MM/DD/YYYY, MM/dd/yyyy
        if (Regex.IsMatch(cell.Style.Numberformat.Format, pattern) && cell.Value is double doubleValue)
        {
            date = DateTime.FromOADate(doubleValue);
            return true;
        }

        date = DateTime.MinValue;
        return false;
    }

    public static string ToCsvDateFormat(this DateTime date)
    {
        return date.ToString(CsvDateFormat);
    }
}