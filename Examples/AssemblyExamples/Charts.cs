using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class Charts : AssemblyExamplesBase
    {
        [TestCase("Bubble chart.docx")]
        [TestCase("Bubble chart.xlsx")]
        [TestCase("Bubble chart.pptx")]
        public void BubbleChart(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:BubbleChart
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template, ArtifactsDir + "Charts.BubbleChart" + extension,
                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
            //ExEnd:BubbleChart
        }

        /// <summary>
        /// Sets colors of chart series point color dynamically based upon expressions.
        /// Feature is supported by version 18.6 or greater.
        /// </summary>
        [TestCase("Dynamic chart point series color.docx")]
        [TestCase("Dynamic chart point series color.xlsx")]
        [TestCase("Dynamic chart point series color.pptx")]
        [TestCase("Dynamic chart point series color.msg")]
        public void DynamicChartSeriesPointColor(string template)
        {
            string extension = Path.GetExtension(template);

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "Charts.DynamicChartSeriesPointColor" + extension,
                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
        }

        /// <summary>
        /// Sets colors of chart series color dynamically based upon expressions.
        /// Feature is supported by version 18.6 or greater.
        /// </summary>
        [TestCase("Dynamic chart series color.docx", "Managers.docx")]
        [TestCase("Dynamic chart series color.xlsx", "Managers.docx")]
        [TestCase("Dynamic chart series color.pptx", "Managers.pptx")]
        [TestCase("Dynamic chart series color.msg", "Managers.docx")]
        public void DynamicChartSeriesColor(string template, string data)
        {
            string extension = Path.GetExtension(template);

            DocumentTable table = new DocumentTable(DataSourcesDir + data, 1,
                new DocumentTableOptions {FirstRowContainsColumnNames = true });
            // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
            Assert.AreEqual(typeof(string), table.Columns["Total_Contract_Price"].Type);
            // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
            // such as summing in templates.
            table.Columns["Total_Contract_Price"].Type = typeof(double);

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template, ArtifactsDir + "Charts.DynamicChartSeriesColor" + extension,
                new DataSourceInfo(table, "managers"), new DataSourceInfo("red", "color"));
        }

        [TestCase("Chart with filtering and dynamic title.docx")]
        [TestCase("Chart with filtering and dynamic title.xlsx")]
        [TestCase("Chart with filtering and dynamic title.pptx")]
        public void FilteringAndDynamicTitle(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:FilteringAndDynamicTitle
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + template, ArtifactsDir + "Charts.FilteringAndDynamicTitle" + extension,
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"),
                new DataSourceInfo("Total Order Quantity by Quarters", "title"));
            //ExEnd:FilteringAndDynamicTitle
        }

        [Test]
        public void BackgroundColor()
        {
            //ExStart:BackgroundColor
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + "Background color.docx",
                ArtifactsDir + "Charts.BackgroundColor.docx",
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), new DataSourceInfo("red", "color"));
            //ExEnd:BackgroundColor
        }

        [Test]
        public void RemoveSelectiveChartSeries()
        {
            //ExStart:RemoveSelectiveChartSeries
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + "Remove selective chart series.docx",
                ArtifactsDir + "Charts.RemoveSelectiveChartSeries.docx",
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"),
                new DataSourceInfo(2, "mode"));
            //ExEnd:RemoveSelectiveChartSeries
        }
    }
}
