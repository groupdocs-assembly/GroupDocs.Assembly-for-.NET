using System;
using System.IO;
using GroupDocs.Assembly;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    public static class RemoveEmptyParagraphs
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # RemoveEmptyParagraphs : Remove empty paragraphs from assembled document using options \n");

            string[] templates = { "Empty paragraph.docx", "Empty paragraph.pptx", "Empty paragraph.msg" };

            foreach (string template in templates)
            {
                string extension = Path.GetExtension(template);

                DocumentAssembler assembler = new DocumentAssembler { Options = DocumentAssemblyOptions.RemoveEmptyParagraphs };

                assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, template),
                    Path.Combine(Constants.OutputPath, "RemoveEmptyParagraphs" + extension),
                    new DataSourceInfo("dummy", "dummy"));
            }
        }
    }
}


