using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    public static class ChangeTargetFileFormat
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # ChangeTargetFileFormat : Change target file format during document assembly (e.g., export to PDF) \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Bubble chart.docx"),
                Path.Combine(Constants.OutputPath, "ChangeTargetFileFormat.pdf"),
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Bubble chart.docx"),
                Path.Combine(Constants.OutputPath, "ChangeTargetFileFormatUsingExplicitSpecifying.pdf"),
                new LoadSaveOptions(FileFormat.Pdf),
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
        }
    }
}


