using System;
using System.IO;
using GroupDocs.Assembly;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    public static class UseMarkdownTemplate
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # UseMarkdownTemplate : Assemble document from Markdown template \n");

            string[] templates = { "Readme.md", "List.md", "Ordered list.md" };

            const string description =
                "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                "office and email file formats based upon template documents and data obtained from various sources " +
                "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

            foreach (string template in templates)
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    Path.Combine(Constants.TemplatesPath, template),
                    Path.Combine(Constants.OutputPath, $"UseMarkdownTemplate-{template}.docx"),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"),
                    new DataSourceInfo(description, "description"));
            }
        }
    }
}


