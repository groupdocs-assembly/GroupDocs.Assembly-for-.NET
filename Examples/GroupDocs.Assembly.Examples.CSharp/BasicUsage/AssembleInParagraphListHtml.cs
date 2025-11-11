using System;
using System.IO;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    using GroupDocs.Assembly;
    using GroupDocs.Assembly.Examples.CSharp.Data;

    public static class AssembleInParagraphListHtml
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleInParagraphListHtml : Assemble HTML document with in-paragraph list from JSON data source \n");

            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "In-paragraph list.html"),
                Path.Combine(Constants.OutputPath, "AssembleInParagraphList.html"),
                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
        }
    }
}


