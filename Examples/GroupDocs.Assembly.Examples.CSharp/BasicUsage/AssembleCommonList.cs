using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleCommonList
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleCommonList : Assemble document with common list from data source \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "In-table master detail.html"),
                Path.Combine(Constants.OutputPath, "AssembleCommonList.html"),
                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
        }
    }
}

