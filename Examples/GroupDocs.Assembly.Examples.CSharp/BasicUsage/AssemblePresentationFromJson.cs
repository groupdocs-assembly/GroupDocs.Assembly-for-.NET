using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssemblePresentationFromJson
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssemblePresentationFromJson : Assemble presentation from JSON data source \n");

            // Do not extract column names from the first row, so that the first row to be treated as a data row.
            // Limit the largest row index, so that only the first four data rows to be loaded.
            DocumentTable table = new DocumentTable(Path.Combine(Constants.DataSourcesPath, "Managers.pptx"), 1,
                new DocumentTableOptions { MaxRowIndex = 3 });

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Table row.pptx"),
                Path.Combine(Constants.OutputPath, "AssemblePresentationFromJson.pptx"),
                new DataSourceInfo(table, "table"));
        }
    }
}


