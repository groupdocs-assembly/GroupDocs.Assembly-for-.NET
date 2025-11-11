using System;
using System.Runtime.InteropServices;

namespace GroupDocs.Assembly.Examples.CSharp.Framework
{
    using GroupDocs.Assembly.Examples.CSharp;
    using GroupDocs.Assembly.Examples.CSharp.QuickStart;
    using GroupDocs.Assembly.Examples.CSharp.BasicUsage;
    using GroupDocs.Assembly.Examples.CSharp.AdvancedUsage;

    internal class RunExamples
    {
        static void Main(string[] args)
        {
            // Get .NET Framework version dynamically
            string frameworkVersion = GetFrameworkVersion();
            
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
            
            AssembleSpreadsheetFromJson.Run();
            AssemblePresentationFromJson.Run();
            AssembleInParagraphListHtml.Run();
            
            // Advanced Usage
            InsertImageDynamically.Run();
            MultipleDataSources.Run();
            UseMarkdownTemplate.Run();
            RemoveEmptyParagraphs.Run();
            ChangeTargetFileFormat.Run();

            Console.WriteLine();
            Console.WriteLine("All done.");
            Console.ReadKey();
        }

        private static string GetFrameworkVersion()
        {
            try
            {                
                return RuntimeInformation.FrameworkDescription;
            }
            catch
            {                
                Version version = Environment.Version;
                return $".NET Framework {version.Major}.{version.Minor}.{version.Build}";
            }
        }
    }
}