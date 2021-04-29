using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class Tables : AssemblyExamplesBase
    {
        /// <summary>
        /// Working With Table Row DataBands in Email Message.
        /// Feature is supported by version 18.2 or greater.
        /// </summary>
        [TestCase("Data table.docx")]
        [TestCase("Data table.xlsx")]
        [TestCase("Data table.pptx")]
        [TestCase("Data table.msg")]
        public void DataTable(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:DataTable
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template, ArtifactsDir + "Tables.DataTable" + extension,
                new DataSourceInfo(DataLayer.ExcelData(), "ds"));
            //ExEnd:DataTable
        }

        /// <summary>
        /// Table Cell Merging in Word Processing documents
        /// features is supported by version 19.1 or greater.
        /// </summary>
        [TestCase("Merging table cells dynamically.docx")]
        [TestCase("Merging table cells dynamically.xlsx")]
        [TestCase("Merging table cells dynamically.pptx")]
        public void CellsMerging(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:TableCellsMerging
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + $"Tables.CellsMerging_{extension}.pdf", new LoadSaveOptions(FileFormat.Pdf),
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            //ExEnd:TableCellsMerging
        }
    }
}
