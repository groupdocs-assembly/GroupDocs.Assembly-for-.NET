using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.AssemblyExamples.BusinessLayer;
using GroupDocs.Assembly;

namespace GroupDocs.AssemblyExamples
{
    public class Program
    {
        static void Main(string[] args)
        {
            //ExStart:ApplyingLicense
            /**
             *  Applying product license
             *  Please uncomment the statement if you do have license.
             */
            //CommonUtilities.ApplyLicense();
            //ExEnd:ApplyingLicense

            #region Generating Bubble Chart Report
            //Generate a bubble chart report in document processing format
            //GenerateReport.GenerateBubbleChart("document", false, false, false, true);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBubbleChart("spreadsheet", false, false, false, true);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBubbleChart("presentation", false, false, false, true);

            //Generate a Bubble chart report in email format
            //GenerateReport.GenerateBubbleChart("email", false, false, false, false);

            #endregion

            #region Generating Bulleted List Report
            //Generate a Bulleted List report in document processing format
            //GenerateReport.GenerateBulletedList("document", false, false, true, false);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBulletedList("spreadsheet", false, false, true, false);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBulletedList("presentation", false, false, true, false);

            //Generate a Bulleted List report in html format
            //GenerateReport.GenerateBulletedList("html", false, false, false, false);

            //Generate a Bulleted List report in text format
            //GenerateReport.GenerateBulletedList("text", false, false, false, false);

            //Generate a Bulleted List report in email format
            //GenerateReport.GenerateBulletedList("email", false, false, false, false);
            #endregion

            #region Generating Chart report with Filtering, Grouping, and Ordering
            //Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document", false, false, true, false);

            //Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet", false, false, false, true);

            //Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation", false, false, false, true);

            //Generate a Chart report with Filtering, Grouping, and Ordering in email format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("email", false, false, false, false);
            #endregion

            #region Generating Common List Report
            //Generate a Common List Report in document processing format
            //GenerateReport.GenerateCommonList("document", false, false, true, false);

            //Generate a Common List Report in spreadsheet format
            //GenerateReport.GenerateCommonList("spreadsheet", false, false, true, false);

            //Generate a Common List Report in presentation format
            //GenerateReport.GenerateCommonList("presentation", false, false, true, false);

            //Generate a Common List Report in html format
            //GenerateReport.GenerateCommonList("html", false, false, false, false);

            //Generate a Common List Report in text format
            //GenerateReport.GenerateCommonList("text", false, false, false, false);

            //Generate a Common List Report in email format
            //GenerateReport.GenerateCommonList("email", false, false, false, false);
            #endregion

            #region Generating Common Master-Detail Report
            //Generate a Common Master-Detail Report in document processing format
            //GenerateReport.GenerateCommonMasterDetail("document", false, false, false, true);

            //Generate a Common Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateCommonMasterDetail("spreadsheet", false, false, false, true);

            //Generate a Common Master-Detail Report in presentation format
            //GenerateReport.GenerateCommonMasterDetail("presentation", false, false, false, true);

            //Generate a Common Master-Detail Report in html format
            //GenerateReport.GenerateCommonMasterDetail("html", false, false, false, false);

            //Generate a Common Master-Detail Report in text format
            //GenerateReport.GenerateCommonMasterDetail("text", false, false, false, false);

            //Generate a Common Master-Detail Report in email format
            //GenerateReport.GenerateCommonMasterDetail("email", false, false, false, false);

            #endregion

            #region Generating In-Paragraph List Report
            //Generate an In-Paragraph List Report in document processing format
            //GenerateReport.GenerateInParagraphList("document", false, false, false, true);

            //Generate an In-Paragraph List Report in spreadsheet format
            //GenerateReport.GenerateInParagraphList("spreadsheet", false, false, false, true);

            //Generate an In-Paragraph List Report in presentation format
            //GenerateReport.GenerateInParagraphList("presentation", false, false, false, true);

            //Generate an In-Paragraph List Report in html format
            //GenerateReport.GenerateInParagraphList("html", false, false, false, false);

            //Generate an In-Paragraph List Report in text format
            //GenerateReport.GenerateInParagraphList("text", false, false, false, false);

            //Generate an In-Paragraph List Report in email format
            //GenerateReport.GenerateInParagraphList("email", false, false, false, false);
            #endregion

            #region Generating In-Table with Alternate Content Report
            //Generate an In-Table List with Alternate Content Report in document processing format
            //GenerateReport.GenerateInTableListWithAlternateContent("document", false, false, false, true);

            //Generate an In-Table List with Alternate Content Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet", false, false, false, true);

            //Generate an In-Table List with Alternate Content Report in presentation format
            //GenerateReport.GenerateInTableListWithAlternateContent("presentation", false, false, false, true);

            //Generate an In-Table List with Alternate Content Report in html format
            //GenerateReport.GenerateInTableListWithAlternateContent("html", false, false, false, false);

            //Generate an In-Table List with Alternate Content Report in email format
            //GenerateReport.GenerateInTableListWithAlternateContent("email", false, false, false, false);

            #endregion

            #region Generating In-Table with Filtering, Grouping and Ordering Report
            //Generate an In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document", false, false, false, true);

            //Generate an In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet", false, false, false, true);

            //Generate an In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation", false, false, false, true);

            //Generate an In-Table List with Filtering, Grouping, and Ordering Report in html format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("html", false, false, false, false);

            //Generate an In-Table List with Filtering, Grouping, and Ordering Report in email format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("email", false, false, false, false);
            #endregion

            #region Generating In-Table List with Highlighted Rows Report
            //Generate an In-Table List with Highlighted Rows Report in document processing format
            //GenerateReport.GenerateInTableListWithHighlightedRows("document", false, false, false, true);

            //Generate an In-Table List with Highlighted Rows Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet", false, false, false, true);

            //Generate an In-Table List with Highlighted Rows Report in presentation format
            //GenerateReport.GenerateInTableListWithHighlightedRows("presentation", false, false, false, true);

            //Generate an In-Table List with Highlighted Rows Report in html format
            //GenerateReport.GenerateInTableListWithHighlightedRows("html", false, false, false, false);

            //Generate an In-Table List with Highlighted Rows Report in email format
            //GenerateReport.GenerateInTableListWithHighlightedRows("email", false, false, false, false);


            #endregion

            #region Generating In-Table List Report
            //Generate an In-Table List Report in document processing format
            //GenerateReport.GenerateInTableList("document", false, false, false, true);

            //Generate an In-Table List Report in spreadsheet format
            //GenerateReport.GenerateInTableList("spreadsheet", false, false, false, true);

            //Generate an In-Table List Report in presentation format
            //GenerateReport.GenerateInTableList("presentation", false, false, false, true);

            //Generate an In-Table List Report in html format
            //GenerateReport.GenerateInTableList("html", false, false, false, false);

            //Generate an In-Table List Report in email format
            //GenerateReport.GenerateInTableList("email", false, false, false, false);
            #endregion

            #region Generating In-Table Master-Detail Report
            //Generate an In-Table Master-Detail Report in document processing format
            //GenerateReport.GenerateInTableMasterDetail("document", false, false, false, true);

            //Generate an In-Table Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateInTableMasterDetail("spreadsheet", false, false, false, true);

            //Generate an In-Table Master-Detail Report in presentation format
            //GenerateReport.GenerateInTableMasterDetail("presentation", false, false, false, true);

            //Generate an In-Table Master-Detail Report in html format
            //GenerateReport.GenerateInTableMasterDetail("html", false, false, false, false);

            //Generate an In-Table Master-Detail Report in email format
            //GenerateReport.GenerateInTableMasterDetail("email", false, false, false, false);
            #endregion

            #region Generating In-Table with Running (Progressive) Total Report
            //Generate an In-Table List with Running (Progressive) Total Report in document processing format
            //GenerateReport.GenerateInTableListWithProgressiveTotal("document", false, false, false, false);

            //Generate an In-Table List with Running (Progressive) Total Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithProgressiveTotal("spreadsheet", false, false, false, false);

            //Generate an In-Table List with Running (Progressive) Total Report in presentation format
            //GenerateReport.GenerateInTableListWithProgressiveTotal("presentation", false, false, false, false);

            //Generate an In-Table List with Running (Progressive) Total Report in html format
            //GenerateReport.GenerateInTableListWithProgressiveTotal("html", false, false, false, false);

            //Generate an In-Table List with Running (Progressive) Total Report in email format
            //GenerateReport.GenerateInTableListWithProgressiveTotal("email", false, false, false, false);
            #endregion

            #region Generating Multicolored Number List Report
            //Generate a Multicolored Numbered List Report in document processing format
            //GenerateReport.GenerateMulticoloredNumberedList("document", false, false, false, true);

            //Generate a Multicolored Numbered List Report in spreadsheet format
            //GenerateReport.GenerateMulticoloredNumberedList("spreadsheet", false, false, false, true);

            //Generate a Multicolored Numbered List Report in presentation format
            //GenerateReport.GenerateMulticoloredNumberedList("presentation", false, false, false, true);

            //Generate a Multicolored Numbered List Report in html format
            //GenerateReport.GenerateMulticoloredNumberedList("html", false, false, false, false);

            //Generate a Multicolored Numbered List Report in email format
            //GenerateReport.GenerateMulticoloredNumberedList("email", false, false, false, false);
            #endregion

            #region Generating Numbered List Report
            //Generate a Numbered List Report in document processing format
            //GenerateReport.GenerateNumberedList("document", false, false, false, true);

            //Generate a Numbered List Report in spreadsheet format
            //GenerateReport.GenerateNumberedList("spreadsheet", false, false, false, true);

            //Generate a Numbered List Report in presentation format
            //GenerateReport.GenerateNumberedList("presentation", false, false, false, true);

            //Generate a Numbered List Report in html format
            //GenerateReport.GenerateNumberedList("html", false, false, false, false);

            //Generate a Numbered List Report in text format
            //GenerateReport.GenerateNumberedList("text", false, false, false, false);

            //Generate a Numbered List Report in email format
            //GenerateReport.GenerateNumberedList("email", false, false, false, false);
            #endregion

            #region Generating Pie Chart Report
            //Generate a Pie Chart Report in document processing format
            //GenerateReport.GeneratePieChart("document", false, false, false, true);

            //Generate a Pie Chart Report in spreadsheet format
            //GenerateReport.GeneratePieChart("spreadsheet", false, false, false, true);

            //Generate a Pie Chart Report in presentation format
            //GenerateReport.GeneratePieChart("presentation", false, false, true, false);

            //Generate a Pie Chart Report in email format
            //GenerateReport.GeneratePieChart("email", false, false, false, false);
            #endregion

            #region Generating Scatter Chart Report
            //Generate a Scatter Chart Report in document processing format
            //GenerateReport.GenerateScatterChart("document", false, false, false, true);

            //Generate a Scatter Chart Report in spreadsheet format
            //GenerateReport.GenerateScatterChart("spreadsheet", false, false, false, true);

            //Generate a Scatter Chart Report in presentation format
            //GenerateReport.GenerateScatterChart("presentation", false, false, false, true);

            //Generate a Scatter Chart Report in email format
            //GenerateReport.GenerateScatterChart("email", false, false, false, false);
            #endregion

            #region Generating Single Row Report
            //Generate a Single Row Report in document processing format
            //GenerateReport.GenerateSingleRow("document", false, false, false, false);

            //Generate a Single Row Report in spreadsheet format
            //GenerateReport.GenerateSingleRow("spreadsheet", false, false, false, true);

            //Generate a Single Row Report in presentation format
            //GenerateReport.GenerateSingleRow("presentation", false, false, false, true);

            //Generate a Single Row Report in html format
            //GenerateReport.GenerateSingleRow("html", false, false, false, false);

            //Generate a Single Row Report in text format
            //GenerateReport.GenerateSingleRow("text", false, false, false, false);

            //Generate a Single Row Report in email format
            //GenerateReport.GenerateSingleRow("email", false, false, false, false);
            #endregion

            #region Generating Report by Recursively and Lazily Accessing the Data
            //GenerateReport.GenerateReportLazilyAndRecursively();
            #endregion

            #region Generating Report using Multiple DataSources
            //Generate a report using multiple data sources in document processing format
            //GenerateReport.GenerateReportUsingMultipleDS("document");
            //Generate a report using multiple data sources in spreadsheet format
            //GenerateReport.GenerateReportUsingMultipleDS("spreadsheet");
            //Generate a report using multiple data sources in presentation format
            //GenerateReport.GenerateReportUsingMultipleDS("presentation");
            #endregion

            #region Template Syntax Formatting

            //Generate document processing formatted reports with desired string or numeric format
            //GenerateReport.TemplateSyntaxFormatting();

            #endregion

            #region Insert Outer Documents

            //Outer document insertion in a report
            //GenerateReport.OuterDocumentInsertion();

            #endregion

            #region Barcode Insertion 

            //add barcode in word processing documents
            //GenerateReport.AddBarCodes("document");
            //add barcode in spreadsheet documents
            //GenerateReport.AddBarCodes("spreadsheet");
            //add barcode in persentation documents 
            //GenerateReport.AddBarCodes("presentation");

            #endregion

            #region Ability to remove selective chart series
            GenerateReport.RemoveSelectiveChartSeries();
            #endregion

            #region Dynamic Chart Axis Title 
            //GenerateReport.DynamicChartAxisTitle();
            #endregion

            #region Dynamic Color in wordpressing document 
            //GenerateReport.DynamicColor();
            #endregion

            //Update fields/formulas in word processing or spreadsheet documents
            //GenerateReport.UpdateWordDocFields("spreadsheet");
            //Use of Next keyword in template syntax
            //GenerateReport.NextIteration();

            //Generate report from excel data source 
            //GenerateReport.UseSpreadsheetAsDataSource();

            //Generate report from presentation data source
            //GenerateReport.UsePresentationTableAsDataSource();

            //Importing word processing table into presentation
            //GenerateReport.ImportingWordProcessingTableIntoPresentation();

            //Importing spread table into html
            //GenerateReport.ImportingSpreadsheetIntoHtmlDocument();

            //Load document table set using default options
            //GenerateReport.LoadDocTableSet("Multiple Tables Data.docx");
            //Load document table set using custom options
            //GenerateReport.LoadDocTableSetWithCustomOptions("Multiple Tables Data.docx");
            //Using DocumentTableSet as Data Source
            //GenerateReport.UseDocumentTableSetAsDataSource("Multiple Tables Data.docx", "Using Document Table Set as Data Source.pptx");
            //GenerateReport.DefiningDocumentTableRelations("Related Tables Data.xlsx", "Using Document Table Relations.docx");
            //GenerateReport.ChangingDocumentTableColumnType("Presentation Templates/Changing Document Table Column Type.pptx");
            Console.WriteLine();
        }
    }
}