using System;
using System.IO;
using GroupDocs.Assembly;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    public static class InsertImageDynamically
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # InsertImageDynamically : Insert images dynamically into document from data source \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Inserting image.docx"),
                Path.Combine(Constants.OutputPath, "InsertImageDynamically.docx"),
                new DataSourceInfo(Path.Combine(Constants.ImagesPath, "no-photo.jpg"), "expression"));
        }
    }
}


