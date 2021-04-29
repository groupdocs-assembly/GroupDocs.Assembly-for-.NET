using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class DynamicallyInsertion : AssemblyExamplesBase
    {
        /// <summary>
        /// Insert Hyperlink Dynamically in Email Document.
        /// Feature is supported by version 18.7 or greater.
        /// </summary>
        [TestCase("Inserting hyperlink.docx")]
        [TestCase("Inserting hyperlink.pptx")]
        [TestCase("Inserting hyperlink.xlsx")]
        [TestCase("Inserting hyperlink.msg")]
        public void Hyperlink(string template)
        {
            string extension = Path.GetExtension(template);

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "DynamicallyInsertion.Hyperlink" + extension,
                new DataSourceInfo("https://www.groupdocs.com/", "uriExpression"),
                new DataSourceInfo("GroupDocs", "displayTextExpression"));
        }

        /// <summary>
        /// Insert Bookmarks Dynamically in Word Document.
        /// Feature is supported by version 20.1 or greater.
        /// </summary>
        [TestCase("Inserting bookmark.docx")]
        [TestCase("Inserting bookmark.xlsx")]
        public void Bookmark(string template)
        {
            string extension = Path.GetExtension(template);

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "DynamicallyInsertion.Bookmark" + extension,
                new DataSourceInfo("gd_bookmark", "bookmark_expression"),
                new DataSourceInfo("GroupDocs", "displayTextExpression"));
        }

        /// <summary>
        /// Insert Image Dynamically in Word Document.
        /// Features is supported by version 20.3 or greater.
        /// </summary>
        [Test]
        public void Image()
        {
            //ExStart:InsertImageDynamically
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Inserting image.docx",
                ArtifactsDir + "DynamicallyInsertion.Image.docx",
                new DataSourceInfo(ImagesDir + "no-photo.jpg", "expression"));
            //ExEnd:InsertImageDynamically
        }

        /// <summary>
        /// Insert Document Dynamically in Word Document.
        /// Features is supported by version 20.3 or greater.
        /// </summary>
        [TestCase("Inserting document.docx")]
        [TestCase("Inserting document.odt")]
        public void Document(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:InsertDocumentDynamically
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + $"DynamicallyInsertion.Document{extension}",
                new DataSourceInfo(TemplatesDir + "Outer document.docx", "document_expression"),
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            //ExEnd:InsertDocumentDynamically
        }

        /// <summary>
        /// Set the values of drop down list items dynamically in Word document.
        /// Features is supported by version 21.3 or greater.
        /// </summary>
        [TestCase("Dropdown.docx", "DropDown")]
        [TestCase("ComboBox.docx", "ComboBox")]
        public void ComboBoxDropDownValues(string template, string element)
        {
            //ExStart:ComboBoxDropDownValues
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + $"DynamicallyInsertion.ComboBoxDropDownValues.{element}.docx",
                new DataSourceInfo("Green apple", "choice_one"),
                new DataSourceInfo("Yellow banana", "choice_two"),
                new DataSourceInfo("Red cherry", "choice_three"),
                new DataSourceInfo("Apple", "choice_one_display_name"),
                new DataSourceInfo("Banana", "choice_two_display_name"),
                new DataSourceInfo("Cherry", "choice_three_display_name"));
            //ExEnd:ComboBoxDropDownValues
        }

        /// <summary>
        /// Set checkbox value dynamically in Word document.
        /// Features is supported by version 20.3 or greater.
        /// </summary>
        [Test]
        public void Checkbox()
        {
            //ExStart:Checkbox
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "CheckBox.docx",
                ArtifactsDir + "DynamicallyInsertion.Checkbox.docx",
                new DataSourceInfo(true, "conditional_expression"));
            //ExEnd:Checkbox
        }

        /// <summary>
        /// This method inserts nested external documents.
        /// </summary>
        [TestCase("Nested external document.docx")]
        [TestCase("Nested external document.msg")]
        public void NestedExternalDocuments(string template)
        {
            string extension = Path.GetExtension(template);

            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "DynamicallyInsertion.NestedExternalDocuments" + extension,
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
        }
    }
}
