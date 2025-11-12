using System;
using System.Runtime.InteropServices;
using GroupDocs.Assembly.Examples.CSharp;
using GroupDocs.Assembly.Examples.CSharp.QuickStart;
using GroupDocs.Assembly.Examples.CSharp.BasicUsage;
using GroupDocs.Assembly.Examples.CSharp.AdvancedUsage;

// Get .NET version dynamically
string frameworkVersion = GetNetVersion();

Console.WriteLine($"Running GroupDocs.Assembly Examples for {frameworkVersion}:");
Console.WriteLine("=====================================================");
Console.WriteLine();

// Display all constant paths
Console.WriteLine("Configuration:");
Console.WriteLine("-----------------------------------------------------");
Console.WriteLine($"Templates Path : \"{Constants.TemplatesPath}\"");
Console.WriteLine($"Data Sources Path : \"{Constants.DataSourcesPath}\"");
Console.WriteLine($"Images Path : \"{Constants.ImagesPath}\"");
Console.WriteLine($"Output Path : \"{Constants.OutputPath}\"");
Console.WriteLine("=====================================================");
Console.WriteLine();

// Quick Start
SetLicenseFromFile.Run();
HelloWorld.Run();

// Basic Usage
AssembleFromCsv.Run();
AssembleFromObject.Run();
AssembleFromJson.Run();
AssembleFromXml.Run();
AssembleFromStringTemplate.Run();

AssembleSpreadsheetFromJson.Run();
AssemblePresentationFromJson.Run();
AssembleInParagraphListHtml.Run();

AssembleTableReport.Run();
AssembleBulletedList.Run();
AssembleCommonList.Run();
AssembleNumberedList.Run();

// Advanced Usage
InsertImageDynamically.Run();
MultipleDataSources.Run();
UseMarkdownTemplate.Run();
RemoveEmptyParagraphs.Run();
ChangeTargetFileFormat.Run();

Console.WriteLine();
Console.WriteLine("All done.");
Console.ReadKey();

static string GetNetVersion()
{
    try
    {
        return RuntimeInformation.FrameworkDescription;
    }
    catch
    {
        Version version = Environment.Version;
        return $".NET {version.Major}.{version.Minor}.{version.Build}";
    }
}
