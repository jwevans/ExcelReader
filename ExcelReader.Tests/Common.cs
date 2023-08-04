namespace ExcelReader.Tests;

using System.Reflection;

internal static class Common
{
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
