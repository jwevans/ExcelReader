﻿using System.Reflection;

namespace ExcelReader.Tests.Files;

public static class Resources
{
    public const string FinancialSample = "ExcelReader.Tests.Files.Financial Sample.xlsx";
    public const string ExployeeSample = "ExcelReader.Tests.Files.Employee Sample.xlsx";
    public const string MultiSheetSample = "ExcelReader.Tests.Files.MultiSheet Sample.xlsx";

    public static Stream GetFileStream(string resourceName)
    {
        var fileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName)!;

        if (fileStream is null)
        {
            throw new ApplicationException($"Could not find embedded resource {resourceName}");
        }

        return fileStream;
    }
}
