using AssemblyExamples.Data;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    public class EmailReports :AssemblyExamplesBase
    {
        [Test]
        public void BubbleChart()
        {
            //ExStart:BubbleChartEmail
            BusinessObjects.EmailDataSourcesObjects dataSources = DataLayer.EmailDataSourceObject(
                TemplatesDir + "Bubble chart.msg",
                DataLayer.GetOrdersData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("orders");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Bubble chart.msg",
                ArtifactsDir + "EmailReports.BubbleChart.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:BubbleChartEmail
        }

        [Test]
        public void DynamicChartAxisTitle()
        {
            //ExStart:DynamicChartAxisTitleEmail
            BusinessObjects.EmailDataSourcesObjects dataSources = DataLayer.EmailDataSourceObject(
                TemplatesDir + "Chart with filtering and dynamic title.msg",
                DataLayer.GetOrdersData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("orders");

            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(
                TemplatesDir + "Chart with filtering and dynamic title.msg",
                ArtifactsDir + "EmailReports.DynamicChartAxisTitle.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject),
                new DataSourceInfo(dataSources.Title, dataSourcesNames.Title)
            );
            //ExEnd:DynamicChartAxisTitleEmail
        }

        [Test]
        public void BulletedList()
        {
            //ExStart:BulletedListEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "Bulleted list.msg",
                    DataLayer.GetProductsData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("products");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Bulleted list.msg",
                ArtifactsDir + "EmailReports.BulletedList.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:BulletedListEmail
        }

        [Test]
        public void CommonList()
        {
            //ExStart:CommonListEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "Common list.msg",
                    DataLayer.PopulateData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("customers");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Common list.msg",
                ArtifactsDir + "EmailReports.CommonList.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:CommonListEmail
        }

        [Test]
        public void CommonMasterDetail()
        {
            //ExStart:CommonMasterDetailEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "Common master detail.msg",
                    DataLayer.PopulateData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("customers");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Common master detail.msg",
                ArtifactsDir + "EmailReports.CommonMasterDetail.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:CommonMasterDetailEmail
        }

        [Test]
        public void InParagraphList()
        {
            //ExStart:InParagraphListEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "In-paragraph list.msg",
                    DataLayer.GetProductsData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("products");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "In-paragraph list.msg",
                ArtifactsDir + "EmailReports.InParagraphList.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:InParagraphListEmail
        }

        [Test]
        public void InTableListWithFiltering()
        {
            //ExStart:InTableListWithFilteringEmail
            BusinessObjects.EmailDataSourcesObjects dataSources = DataLayer.EmailDataSourceObject(
                TemplatesDir + "In-table list with filtering.msg",
                DataLayer.GetOrdersData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("orders");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + "In-table list with filtering.msg",
                ArtifactsDir + "EmailReports.InTableListWithFiltering.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:InTableListWithFilteringEmail
        }

        [Test]
        public void InTableListWithHighlightedRows()
        {
            //ExStart:InTableListWithHighlightedRowsEmail
            BusinessObjects.EmailDataSourcesObjects dataSources = DataLayer.EmailDataSourceObject(
                TemplatesDir + "In-table list with highlighted rows.msg",
                DataLayer.GetOrdersData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("orders");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "In-table list with highlighted rows.msg",
                ArtifactsDir + "EmailReports.InTableListWithHighlightedRows.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:InTableListWithHighlightedRowsEmail
        }

        [Test]
        public void InTableList()
        {
            //ExStart:InTableListEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "In-table list.msg",
                    DataLayer.PopulateData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("customers");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "In-table list.msg",
                ArtifactsDir + "EmailReports.InTableList.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:InTableListEmail
        }

        [Test]
        public void InTableMasterDetail()
        {
            //ExStart:InTableMasterDetailEmail
            BusinessObjects.EmailDataSourcesObjects dataSources = DataLayer.EmailDataSourceObject(
                TemplatesDir + "In-table master detail.msg", DataLayer.PopulateData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("customers");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "In-table master detail.msg",
                ArtifactsDir + "EmailReports.InTableMasterDetail.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:InTableMasterDetailEmail
        }

        [Test]
        public void InTableListWithTotal()
        {
            //ExStart:InTableListWithTotalEmail
            BusinessObjects.EmailDataSourcesObjects dataSources = DataLayer.EmailDataSourceObject(
                TemplatesDir + "In-table list with total.msg", DataLayer.GetOrdersData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("orders");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + "In-table list with total.msg",
                ArtifactsDir + "EmailReports.InTableListWithTotal.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:InTableListWithTotalEmail
        }

        [Test]
        public void NumberedList()
        {
            //ExStart:NumberedListEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "Numbered list.msg",
                    DataLayer.GetProductsData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("products");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Numbered list.msg",
                ArtifactsDir + "EmailReports.NumberedList.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:NumberedListEmail
        }

        [Test]
        public void TableCellsMerging()
        {
            //ExStart:TableCellsMergingEmail
            BusinessObjects.EmailDataSourcesObjects dataSources =
                DataLayer.EmailDataSourceObject(TemplatesDir + "Merging table cells dynamically.msg",
                    DataLayer.PopulateData());
            BusinessObjects.EmailDataSourcesNames dataSourcesNames = DataLayer.EmailDataSourceName("customers");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Merging table cells dynamically.msg",
                ArtifactsDir + "EmailReports.TableCellsMerging.msg",
                new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
            );
            //ExEnd:TableCellsMergingEmail
        }
    }
}
