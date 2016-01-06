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
            //GenerateReport.GenerateBubbleChart("document", false, false, false);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBubbleChart("spreadsheet", false, false, false);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBubbleChart("presentation", false, false, false);
            #endregion
            
            #region Generating Bulleted List Report
            //Generate a Bulleted List report in document processing format
            //GenerateReport.GenerateBulletedList("document", false, false, false);

            //Generate a Bulleted List report in spreadsheet format
            GenerateReport.GenerateBulletedList("spreadsheet", false, false, true);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBulletedList("presentation", false, false, true);
            #endregion
            
            #region Generating Chart report with Filtering, Grouping, and Ordering
            //Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document", false, false, false);

            //Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet", false, false, false);

            //Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation", false, false, false);
            #endregion
            
            #region Generating Common List Report
            //Generate a Common List Report in document processing format
            //GenerateReport.GenerateCommonList("document", true, false, false);

            //Generate a Common List Report in spreadsheet format
            //GenerateReport.GenerateCommonList("spreadsheet", false, false, false);

            //Generate a Common List Report in presentation format
            //GenerateReport.GenerateCommonList("presentation", false, false, false);
            #endregion
            
            #region Generating Common Master-Detail Report
            //Generate a Common Master-Detail Report in document processing format
            //GenerateReport.GenerateCommonMasterDetail("document", false, false, false);

            //Generate a Common Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateCommonMasterDetail("spreadsheet", false, false, false);

            //Generate a Common Master-Detail Report in presentation format
            //GenerateReport.GenerateCommonMasterDetail("presentation", false, false, false);
            #endregion
            
            #region Generating In-Paragraph List Report
            //Generate a In-Paragraph List Report in document processing format
            //GenerateReport.GenerateInParagraphList("document", false, false, false);

            //Generate a In-Paragraph List Report in spreadsheet format
            //GenerateReport.GenerateInParagraphList("spreadsheet", false, false, true);

            //Generate a In-Paragraph List Report in presentation format
            //GenerateReport.GenerateInParagraphList("presentation", false, true, false);
            #endregion
            
            #region Generating In-Table with Alternate Content Report
            //Generate a In-Table List with Alternate Content Report in document processing format
            //GenerateReport.GenerateInTableListWithAlternateContent("document", false, false, false);

            //Generate a In-Table List with Alternate Content Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet", false, false, true);

            //Generate a In-Table List with Alternate Content Report in presentation format
            //GenerateReport.GenerateInTableListWithAlternateContent("presentation", false, false, true);
            #endregion
           
            #region Generating In-Table with Filtering, Grouping and Ordering Report
            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document", false, false, false);

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet", false, false, true);

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation", false, false, true);
            #endregion
            
            #region Generating In-Table List with Highlighted Rows Report
            //Generate a In-Table List with Highlighted Rows Report in document processing format
            //GenerateReport.GenerateInTableListWithHighlightedRows("document", false, false, false);

            //Generate a In-Table List with Highlighted Rows Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet", false, false, true);

            //Generate a In-Table List with Highlighted Rows Report in presentation format
            //GenerateReport.GenerateInTableListWithHighlightedRows("presentation", false, false, true);
            #endregion
            
            #region Generating In-Table List Report
            //Generate a In-Table List Report in document processing format
            //GenerateReport.GenerateInTableList("document", false, false, false);

            //Generate a In-Table List Report in spreadsheet format
            //GenerateReport.GenerateInTableList("spreadsheet", false, false, true);

            //Generate a In-Table List Report in presentation format
            //GenerateReport.GenerateInTableList("presentation", false, false, true);
            #endregion
            
            #region Generating In-Table Master-Detail Report
            //Generate a In-Table Master-Detail Report in document processing format
            //GenerateReport.GenerateInTableMasterDetail("document", false, false, false);

            //Generate a In-Table Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateInTableMasterDetail("spreadsheet", false, false, true);

            //Generate a In-Table Master-Detail Report in presentation format
            //GenerateReport.GenerateInTableMasterDetail("presentation", false, false, true);
            #endregion
            
            #region Generating Multicolored Number List Report
            //Generate a Multicolored Numbered List Report in document processing format
            //GenerateReport.GenerateMulticoloredNumberedList("document", false, false, false);

            //Generate a Multicolored Numbered List Report in spreadsheet format
            //GenerateReport.GenerateMulticoloredNumberedList("spreadsheet", false, false, true);

            //Generate a Multicolored Numbered List Report in presentation format
            //GenerateReport.GenerateMulticoloredNumberedList("presentation", false, false, true);
            #endregion
            
            #region Generating Numbered List Report
            //Generate a Numbered List Report in document processing format
            //GenerateReport.GenerateNumberedList("document", false, false, false);

            //Generate a Numbered List Report in spreadsheet format
            //GenerateReport.GenerateNumberedList("spreadsheet", false, false, true);

            //Generate a Numbered List Report in presentation format
            //GenerateReport.GenerateNumberedList("presentation", false, false, true);
            #endregion
            
            #region Generating Pie Chart Report
            //Generate a Pie Chart Report in document processing format
            //GenerateReport.GeneratePieChart("document", false, false, false);
            
            //Generate a Pie Chart Report in spreadsheet format
            //GenerateReport.GeneratePieChart("spreadsheet", false, false, false);

            //Generate a Pie Chart Report in presentation format
            //GenerateReport.GeneratePieChart("presentation", false, false, false);
            #endregion
            
            #region Generating Scatter Chart Report
            //Generate a Scatter Chart Report in document processing format
            //GenerateReport.GenerateScatterChart("document", false, false, false);

            //Generate a Scatter Chart Report in spreadsheet format
            //GenerateReport.GenerateScatterChart("spreadsheet", false, false, false);

            //Generate a Scatter Chart Report in presentation format
            //GenerateReport.GenerateScatterChart("presentation", false, false, false);
            #endregion
            
            #region Generating Single Row Report
            //Generate a Single Row Report in document processing format
            //GenerateReport.GenerateSingleRow("document", false, false, false);

            //Generate a Single Row Report in spreadsheet format
            //GenerateReport.GenerateSingleRow("spreadsheet", false, false, true);

            //Generate a Single Row Report in presentation format
            //GenerateReport.GenerateSingleRow("presentation", false, false, false);
            #endregion
            
           
        }
    }
}

