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
            /**
             *  Applying product license
             *  Please uncomment the statement if you do have license.
             */
            CommonUtilities.ApplyLicense();

            #region Generating Bubble Chart Report
            //Generate a bubble chart report in document processing format
            GenerateReport.GenerateBubbleChart("document");

            //Generate a Bulleted List report in spreadsheet format
            GenerateReport.GenerateBubbleChart("spreadsheet");

            //Generate a Bulleted List report in presentation format
            GenerateReport.GenerateBubbleChart("presentation");
            #endregion

            #region Generating Bulleted List Report
            //Generate a Bulleted List report in document processing format
            GenerateReport.GenerateBulletedList("document");

            //Generate a Bulleted List report in spreadsheet format
            GenerateReport.GenerateBulletedList("spreadsheet");

            //Generate a Bulleted List report in presentation format
            GenerateReport.GenerateBulletedList("presentation");
            #endregion

            #region Generating Chart report with Filtering, Grouping, and Ordering
            //Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document");

            //Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet");

            //Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation");
            #endregion

            #region Generating Common List Report
            //Generate a Common List Report in document processing format
            GenerateReport.GenerateCommonList("document");

            //Generate a Common List Report in spreadsheet format
            GenerateReport.GenerateCommonList("spreadsheet");

            //Generate a Common List Report in presentation format
            GenerateReport.GenerateCommonList("presentation");
            #endregion

            #region Generating Common Master-Detail Report
            //Generate a Common Master-Detail Report in document processing format
            GenerateReport.GenerateCommonMasterDetail("document");

            //Generate a Common Master-Detail Report in spreadsheet format
            GenerateReport.GenerateCommonMasterDetail("spreadsheet");

            //Generate a Common Master-Detail Report in presentation format
            GenerateReport.GenerateCommonMasterDetail("presentation");
            #endregion

            #region Generating In-Paragraph List Report
            //Generate a In-Paragraph List Report in document processing format
            GenerateReport.GenerateInParagraphList("document");

            //Generate a In-Paragraph List Report in spreadsheet format
            GenerateReport.GenerateInParagraphList("spreadsheet");

            //Generate a In-Paragraph List Report in presentation format
            GenerateReport.GenerateInParagraphList("presentation");
            #endregion

            #region Generating In-Table with Alternate Content Report
            //Generate a In-Table List with Alternate Content Report in document processing format
            GenerateReport.GenerateInTableListWithAlternateContent("document");

            //Generate a In-Table List with Alternate Content Report in spreadsheet format
            GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet");

            //Generate a In-Table List with Alternate Content Report in presentation format
            GenerateReport.GenerateInTableListWithAlternateContent("presentation");
            #endregion

            #region Generating In-Table with Filtering, Grouping and Ordering Report
            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document");

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet");

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation");
            #endregion

            #region Generating In-Table List with Highlighted Rows Report
            //Generate a In-Table List with Highlighted Rows Report in document processing format
            GenerateReport.GenerateInTableListWithHighlightedRows("document");

            //Generate a In-Table List with Highlighted Rows Report in spreadsheet format
            GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet");

            //Generate a In-Table List with Highlighted Rows Report in presentation format
            GenerateReport.GenerateInTableListWithHighlightedRows("presentation");
            #endregion

            #region Generating In-Table List Report
            //Generate a In-Table List Report in document processing format
            GenerateReport.GenerateInTableList("document");

            //Generate a In-Table List Report in spreadsheet format
            GenerateReport.GenerateInTableList("spreadsheet");

            //Generate a In-Table List Report in presentation format
            GenerateReport.GenerateInTableList("presentation");
            #endregion

            #region Generating In-Table Master-Detail Report
            //Generate a In-Table Master-Detail Report in document processing format
            GenerateReport.GenerateInTableMasterDetail("document");

            //Generate a In-Table Master-Detail Report in spreadsheet format
            GenerateReport.GenerateInTableMasterDetail("spreadsheet");

            //Generate a In-Table Master-Detail Report in presentation format
            GenerateReport.GenerateInTableMasterDetail("presentation");
            #endregion

            #region Generating Multicolored Number List Report
            //Generate a Multicolored Numbered List Report in document processing format
            GenerateReport.GenerateMulticoloredNumberedList("document");

            //Generate a Multicolored Numbered List Report in spreadsheet format
            GenerateReport.GenerateMulticoloredNumberedList("spreadsheet");

            //Generate a Multicolored Numbered List Report in presentation format
            GenerateReport.GenerateMulticoloredNumberedList("presentation");
            #endregion

            #region Generating Numbered List Report
            //Generate a Numbered List Report in document processing format
            GenerateReport.GenerateNumberedList("document");

            //Generate a Numbered List Report in spreadsheet format
            GenerateReport.GenerateNumberedList("spreadsheet");

            //Generate a Numbered List Report in presentation format
            GenerateReport.GenerateNumberedList("presentation");
            #endregion

            #region Generating Pie Chart Report
            //Generate a Pie Chart Report in document processing format
            GenerateReport.GeneratePieChart("document");

            //Generate a Pie Chart Report in spreadsheet format
            GenerateReport.GeneratePieChart("spreadsheet");

            //Generate a Pie Chart Report in presentation format
            GenerateReport.GeneratePieChart("presentation");
            #endregion

            #region Generating Scatter Chart Report
            //Generate a Scatter Chart Report in document processing format
            GenerateReport.GenerateScatterChart("document");

            //Generate a Scatter Chart Report in spreadsheet format
            GenerateReport.GenerateScatterChart("spreadsheet");

            //Generate a Scatter Chart Report in presentation format
            GenerateReport.GenerateScatterChart("presentation");
            #endregion

            #region Generating Single Row Report
            //Generate a Single Row Report in document processing format
            GenerateReport.GenerateSingleRow("document");

            //Generate a Single Row Report in spreadsheet format
            GenerateReport.GenerateSingleRow("spreadsheet");

            //Generate a Single Row Report in presentation format
            GenerateReport.GenerateSingleRow("presentation");
            #endregion

           
        }
    }
}

