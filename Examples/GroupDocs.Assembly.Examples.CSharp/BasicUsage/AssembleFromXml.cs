using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleFromXml
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleFromXml : Assemble document from XML data source \n");

            XmlDataSource dataSource = new XmlDataSource(Path.Combine(Constants.DataSourcesPath, "Managers.xml"));

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Data destination with nested elements.docx"),
                Path.Combine(Constants.OutputPath, "AssembleFromXml.docx"),
                new DataSourceInfo(dataSource, "managers"));
        }
    }
}


