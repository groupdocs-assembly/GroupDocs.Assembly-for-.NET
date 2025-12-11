using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleWithOoxmlCompliance
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleWithOoxmlCompliance : Assemble document with specified OOXML compliance level \n");

            DocumentAssembler assembler = new DocumentAssembler();

            // Create LoadSaveOptions with explicit OOXML compliance
            var options = new LoadSaveOptions(FileFormat.Docx);
            options.OoxmlCompliance = OoxmlCompliance.Strict;

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Table row.docx"),
                Path.Combine(Constants.OutputPath, "AssembleWithOoxmlCompliance.docx"),
                options,
                new DataSourceInfo(DataLayer.ExcelData(), "contracts"));
        }
    }
}

