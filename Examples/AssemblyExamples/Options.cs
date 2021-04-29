using System;
using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class Options : AssemblyExamplesBase
    {
        /// <summary>
        /// Loading of template documents from HTML with resources from an explicitly specified folder.
        /// Features is supported by version 19.5 or greater.
        /// </summary>
        [Test]
        public void LoadExternalResourceFiles()
        {
            //ExStart:ExternalResourceFiles
            DocumentAssembler assembler = new DocumentAssembler();

            LoadSaveOptions loadSaveOptions = new LoadSaveOptions
            {
                // Resolve URIs from the specified alternative folder.
                ResourceLoadBaseUri = TemplatesDir + "ExternalResourceFiles"
            };

            assembler.AssembleDocument(TemplatesDir + "ExternalResourceFiles.htm",
                ArtifactsDir + "Options.ExternalResourceFiles.docx", loadSaveOptions,
                new DataSourceInfo("It should be a sport car image.", "value"));
            //ExEnd:ExternalResourceFiles
        }

        /// <summary>
        /// Saving of external resource files in a specified folder at relative path while saving output to HTML.
        /// Features is supported by version 19.5 or greater.
        /// </summary>
        [Test]
        public void SaveExternalResourceFiles()
        {
            //ExStart:SaveDocToHtmlWithResource
            DocumentAssembler assembler = new DocumentAssembler();

            LoadSaveOptions loadSaveOptions = new LoadSaveOptions
            {
                // Resolve URIs from the specified alternative folder.
                ResourceSaveFolder = ArtifactsDir + "SaveExternalResourceFiles"
            };

            assembler.AssembleDocument(TemplatesDir + "ExternalResourceFiles.docx",
                ArtifactsDir + "Options.SaveExternalResourceFiles.htm", loadSaveOptions,
                new DataSourceInfo("Hello!", "value"));
            //ExEnd:SaveDocToHtmlWithResource
        }

        /// <summary>
        /// In-lining syntax error messages into templates
        /// features is supported by version 19.3 or greater.
        /// </summary>
        [Test]
        public void InLineSyntaxError()
        {
            //ExStart:InLineSyntaxError
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.Options |= DocumentAssemblyOptions.InlineErrorMessages;

            // The AssembleDocument will return a boolean value to indicate the success or failed with inline error.
            Console.WriteLine(assembler.AssembleDocument(TemplatesDir + "Inline error.docx",
                ArtifactsDir + "Options.InLineSyntaxError.pdf", new LoadSaveOptions(FileFormat.Pdf),
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"))
                ? "No error found in template"
                : "Do something with a report containing a template syntax error.");
            //ExEnd:InLineSyntaxError
        }

        [TestCase("Update field.docx")]
        [TestCase("Update field.xlsx")]
        public void UpdateFields(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:UpdateFields
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.Options |= DocumentAssemblyOptions.UpdateFieldsAndFormulas;

            assembler.AssembleDocument(TemplatesDir + template, "Options.UpdateFields" + extension,
                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
            //ExEnd:UpdateFields
        }

        [TestCase("Empty paragraph.docx")]
        [TestCase("Empty paragraph.pptx")]
        [TestCase("Empty paragraph.msg")]
        public void RemoveEmptyParagraphs(string template)
        {
            string extension = Path.GetExtension(template);

            DocumentAssembler assembler = new DocumentAssembler { Options = DocumentAssemblyOptions.RemoveEmptyParagraphs };

            assembler.AssembleDocument(TemplatesDir + template, ArtifactsDir + "Options.RemoveEmptyParagraphs" + extension,
                new DataSourceInfo("dummy", "dummy"));
        }

        [Test]
        public void SyntaxFormatting()
        {
            //ExStart:SyntaxFormatting
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Numeric format.odt",
                ArtifactsDir + "Options.SyntaxFormatting.odt",
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            //ExEnd:SyntaxFormatting 
        }

        /// <summary>
        /// Support for analogue of Microsoft Word NEXT Field.
        /// </summary>
        [Test]
        public void NextField()
        {
            //ExStart:NextField
            DocumentAssembler assembler = new DocumentAssembler();
            // Call AssembleDocument to generate Report.
            assembler.AssembleDocument(TemplatesDir + "Next field.docx",
                ArtifactsDir + "Options.NextField.docx",
                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
            //ExEnd:NextField
        }
    }
}
