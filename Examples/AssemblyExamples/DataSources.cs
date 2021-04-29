using System.IO;
using System.Text;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class DataSources : AssemblyExamplesBase
    {
        /// <summary>
        /// Working with JSON data sources.
        /// Features is supported by version 19.10 or greater.
        /// </summary>
        [Test]
        public void Json()
        {
            //ExStart:JsonDataSource
            JsonDataSource dataSource = new JsonDataSource(DataSourcesDir + "Managers.json");

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(TemplatesDir + "Data destination with nested elements.docx",
                ArtifactsDir + "DataSources.JsonDataSource.docx",
                new DataSourceInfo(dataSource, "managers"));
            //ExEnd:JsonDataSource
        }

        /// <summary>
        /// Working with XML data sources.
        /// Features is supported by version 19.10 or greater.
        /// </summary>
        [Test]
        public void Xml()
        {
            //ExStart:XmlDataSource
            XmlDataSource dataSource = new XmlDataSource(DataSourcesDir + "Managers.xml");

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(TemplatesDir + "Data destination with nested elements.docx",
                ArtifactsDir + "DataSources.XmlDataSource.docx",
                new DataSourceInfo(dataSource, "managers"));
            //ExEnd:XmlDataSource
        }

        /// <summary>
        /// Working with csv data sources.
        /// Features is supported by version 19.10 or greater.
        /// </summary>
        [Test]
        public void Csv()
        {
            //ExStart:CsvDataSource
            CsvDataSource dataSource =
                new CsvDataSource(DataSourcesDir + "Persons.csv", new CsvDataLoadOptions(true));

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Data destination with nested elements.txt",
                ArtifactsDir + "DataSources.CsvDataSource.docx",
                new DataSourceInfo(dataSource, "persons"));
            //ExEnd:CsvDataSource
        }

        [TestCase("Multiple data source.odt")]
        [TestCase("Multiple data source.ods")]
        [TestCase("Multiple data source.odp")]
        public void MultipleDataSources(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:MultipleDataSources
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "DataSources.MultipleDataSources" + extension,
                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"),
                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
            //ExEnd:MultipleDataSources
        }

        /// <summary>
        /// Generate report from excel data source.
        /// </summary>
        [Test]
        public void Spreadsheet()
        {
            //ExStart:SpreadsheetDataSource
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Table row.docx",
                ArtifactsDir + "DataSources.SpreadsheetDataSource.docx",
                new DataSourceInfo(DataLayer.ExcelData(), "contracts"));
            //ExEnd:SpreadsheetDataSource
        }

        /// <summary>
        /// Generate report from presentation data source.
        /// </summary>
        [Test]
        public void Presentation()
        {
            // Do not extract column names from the first row, so that the first row to be treated as a data row.
            // Limit the largest row index, so that only the first four data rows to be loaded.
            DocumentTable table = new DocumentTable(DataSourcesDir + "Managers.pptx", 1,
                new DocumentTableOptions { MaxRowIndex = 3 });

            // Check column count and names.
            Assert.AreEqual(2, table.Columns.Count);
            // NOTE: Default column names are used, because we do not extract the names from the first row.
            Assert.AreEqual("Column1", table.Columns[0].Name);
            Assert.AreEqual("Column2", table.Columns[1].Name);

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Table row.pptx",
                ArtifactsDir + "DataSources.PresentationDataSource.pptx",
                new DataSourceInfo(table, "table"));
        }

        [Test]
        public void String()
        {
            DocumentAssembler assembler = new DocumentAssembler();

            const string sourceString = @"<<[yourValue]>>";
            byte[] sourceBytes = Encoding.UTF8.GetBytes(sourceString);
            byte[] targetBytes;

            using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
            {
                using (MemoryStream targetStream = new MemoryStream())
                {
                    assembler.AssembleDocument(sourceStream, targetStream,
                        new DataSourceInfo("Hello, World!", "yourValue"));
                    targetBytes = targetStream.ToArray();
                }
            }

            string targetString = Encoding.UTF8.GetString(targetBytes);

            Assert.AreEqual("Hello, World!", targetString);
        }
    }
}
