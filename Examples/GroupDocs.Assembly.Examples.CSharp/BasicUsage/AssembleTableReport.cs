using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleTableReport
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleTableReport : Assemble document with table report from data source \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Data table.docx"),
                Path.Combine(Constants.OutputPath, "AssembleTableReport.docx"),
                new DataSourceInfo(DataLayer.ExcelData(), "ds"));
        }
    }
}

