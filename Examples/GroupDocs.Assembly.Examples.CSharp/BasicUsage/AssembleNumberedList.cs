using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleNumberedList
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleNumberedList : Assemble document with numbered list from data source \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Numbered list.txt"),
                Path.Combine(Constants.OutputPath, "AssembleNumberedList.docx"),
                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
        }
    }
}

