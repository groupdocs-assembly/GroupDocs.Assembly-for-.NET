using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleFromJson
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleFromJson : Assemble document from JSON data source \n");

            JsonDataSource dataSource = new JsonDataSource(Path.Combine(Constants.DataSourcesPath, "Managers.json"));

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Data destination with nested elements.docx"),
                Path.Combine(Constants.OutputPath, "AssembleFromJson.docx"),
                new DataSourceInfo(dataSource, "managers"));
        }
    }
}


