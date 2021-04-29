using System.IO;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class Markdown : AssemblyExamplesBase
    {
        /// <summary>
        /// Saving an assembled Markdown document to a Word Processing format using file extension.
        /// Features is supported by version 19.8 or greater.
        /// </summary>
        [TestCase("Readme.md")]
        [TestCase("List.md")]
        [TestCase("Ordered list.md")]
        public void MdToWord(string template)
        {
            //ExStart:MdToWord
            const string description =
                "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                "office and email file formats based upon template documents and data obtained from various sources " +
                "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + template,
                ArtifactsDir + $"Markdown.MdToWord-{template}.docx",
                new DataSourceInfo("GroupDocs.Assembly for .NET", "product"),
                new DataSourceInfo(description, "description"));
            //ExEnd:MdToWord
        }

        /// <summary>
        /// Saving an assembled Word Processing document or email to Markdown using file extension.
        /// Features is supported by version 19.8 or greater.
        /// </summary>
        [Test]
        public void WordToMd()
        {
            //ExStart:WordToMd
            const string description =
                "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                "office and email file formats based upon template documents and data obtained from various sources " +
                "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Readme.docx",
                ArtifactsDir + "Markdown.WordToMd.md",
                new DataSourceInfo("GroupDocs.Assembly for .NET", "product"),
                new DataSourceInfo(description, "description"));
            //ExEnd:WordToMd
        }

        /// <summary>
        /// Saving an assembled Word Processing document or email to Markdown using explicit specifying.
        /// Features is supported by version 19.8 or greater.
        /// </summary>
        [Test]
        public void WordToMdStream()
        {
            //ExStart:WordToMdStream
            const string description =
                "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                "office and email file formats based upon template documents and data obtained from various sources " +
                "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

            using (Stream sourceStream = File.OpenRead(TemplatesDir + "ReadMe.docx"))
            {
                using (FileStream targetStream = File.Create(ArtifactsDir + "Markdown.WordToMdStream.md"))
                {
                    DocumentAssembler assembler = new DocumentAssembler();
                    assembler.AssembleDocument(sourceStream, targetStream, new LoadSaveOptions(FileFormat.Markdown),
                        new DataSourceInfo("GroupDocs.Assembly for .NET(Using Stream)", "product"),
                        new DataSourceInfo(description, "description"));
                }
            }
            //ExEnd:WordToMdStream
        }

        /// <summary>
        /// Saving Markdown tables to Word document.
        /// Feature is supported by version 20.9 or greater.
        /// </summary>
        [Test]
        public void Tables()
        {
            //ExStart:MarkdownTables
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Tables.md",
                ArtifactsDir + "Markdown.Tables.docx",
                new DataSourceInfo("Lettuce", "product_1_name"),
                new DataSourceInfo("Carrot", "product_2_name"),
                new DataSourceInfo("35", "product_1_quantity"),
                new DataSourceInfo("47", "product_2_quantity"));
            //ExEnd:MarkdownTables
        }

        /// <summary>
        /// Saving Markdown Autolinks to Word document.
        /// Feature is supported by version 20.9 or greater.
        /// </summary>
        [Test]
        public void Autolinks()
        {
            //ExStart:MarkdownAutolinks
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Autolinks.md",
                ArtifactsDir + "Markdown.Autolinks.docx",
                new DataSourceInfo("<https://forum.aspose.com/>", "url"));
            //ExEnd:MarkdownAutolinks
        }

        /// <summary>
        /// Saving Markdown inline links to Word document.
        /// Feature is supported by version 20.11 or greater.
        /// </summary>
        [Test]
        public void InlineLinks()
        {
            //ExStart:MarkdownInlineLinks
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Inline links.md",
                ArtifactsDir + "Markdown.InlineLinks.docx",
                new DataSourceInfo("Aspose forums", "link_text"),
                new DataSourceInfo("https://forum.aspose.com/", "link"));
            //ExEnd:MarkdownInlineLinks
        }

        /// <summary>
        /// Saving Markdown inline images to Word document.
        /// Feature is supported by version 20.11 or greater.
        /// </summary>
        [Test]
        public void InlineImages()
        {
            //ExStart:MarkdownInlineImages
            // Ensure that the relative image URI points to an existing image when in the output document.
            string imgDirectory = ArtifactsDir + "Images";

            if (!Directory.Exists(imgDirectory))
                Directory.CreateDirectory(imgDirectory);

            File.Copy(ImagesDir + "Logo.jpg",
                ArtifactsDir + "Images\\Logo.jpg");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Inline image.md",
                ArtifactsDir + "Markdown.InlineImages.docx",
                new DataSourceInfo("Aspose Logo", "alt_text"),
                new DataSourceInfo("https://docs.aspose.com/images/Aspose-image-for-open-graph.jpg", "image_URI"));
            //ExEnd:MarkdownInlineImages
        }
    }
}
