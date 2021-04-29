using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class Lists : AssemblyExamplesBase
    {
        [TestCase("Bulleted list.odt")]
        [TestCase("Bulleted list.ods")]
        [TestCase("Bulleted list.odp")]
        [TestCase("Bulleted list.html")]
        [TestCase("Bulleted list.txt")]
        public void BulletedList(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:BulletedList
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "Lists.BulletedList" + extension,
                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
            //ExEnd:BulletedList
        }

        [TestCase("Common list.odt")]
        [TestCase("Common list.ods")]
        [TestCase("Common list.odp")]
        [TestCase("Common list.html")]
        [TestCase("Common list.txt")]
        public void CommonList(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:CommonList
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "Lists.CommonList" + extension,
                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
            //ExEnd:CommonList
        }

        [TestCase("In-paragraph list.odt")]
        [TestCase("In-paragraph list.ods")]
        [TestCase("In-paragraph list.odp")]
        [TestCase("In-paragraph list.html")]
        [TestCase("In-paragraph list.txt")]
        public void InParagraphList(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:InParagraphList
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "Lists.InParagraphList" + extension,
                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
            //ExEnd:InParagraphList
        }

        [TestCase("In-table list with filtering.odt")]
        [TestCase("In-table list with highlighted rows.odt")]
        [TestCase("In-table list with total.docx")]
        [TestCase("In-table list with filtering.ods")]
        [TestCase("In-table list with highlighted rows.ods")]
        [TestCase("In-table list with total.xlsx")]
        [TestCase("In-table list with filtering.odp")]
        [TestCase("In-table list with highlighted rows.odp")]
        [TestCase("In-table list with total.pptx")]
        [TestCase("In-table list with filtering.html")]
        [TestCase("In-table list with highlighted rows.html")]
        [TestCase("In-table list with total.html")]
        public void InTableList(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:InTableList
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(
                TemplatesDir + template,
                ArtifactsDir + "Lists.InTableList" + extension,
                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
            //ExEnd:InTableList
        }

        [TestCase("In-table master detail.odt")]
        [TestCase("In-table master detail.ods")]
        [TestCase("In-table master detail.odp")]
        [TestCase("In-table master detail.html")]
        public void InTableListMasterDetail(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:InTableListMasterDetail
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(
                TemplatesDir + template,
                ArtifactsDir + "Lists.InTableListMasterDetail" + extension,
                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
            //ExEnd:InTableListMasterDetail
        }

        [TestCase("Numbered list.odt")]
        [TestCase("Numbered list.ods")]
        [TestCase("Numbered list.odp")]
        [TestCase("Numbered list.html")]
        [TestCase("Numbered list.txt")]
        public void NumberedList(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:NumberedList
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "Lists.NumberedList" + extension,
                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
            //ExEnd:NumberedList
        }

        [TestCase("Numbered list with RestartNum.docx")]
        [TestCase("Numbered list with RestartNum.msg")]
        public void NumberedListRestartNum(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:NumberedListRestartNum
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "Lists.NumberedListRestartNum" + extension,
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            //ExEnd:NumberedListRestartNum
        }
    }
}
