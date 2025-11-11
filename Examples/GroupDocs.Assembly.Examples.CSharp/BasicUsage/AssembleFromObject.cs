using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    
    public static class AssembleFromObject
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleFromObject : Assemble document from object data source \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Table row.docx"),
                Path.Combine(Constants.OutputPath, "AssembleFromObject.docx"),
                new DataSourceInfo(DataLayer.ExcelData(), "contracts"));
        }
    }
}


