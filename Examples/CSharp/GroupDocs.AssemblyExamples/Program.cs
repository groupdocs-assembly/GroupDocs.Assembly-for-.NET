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
            GenerateReport.GenerateBubbleChart("document", false, true);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBubbleChart("spreadsheet", false, true);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBubbleChart("presentation", false, true);
            #endregion
            
            #region Generating Bulleted List Report
            //Generate a Bulleted List report in document processing format
            //GenerateReport.GenerateBulletedList("document", false, true);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBulletedList("spreadsheet", false, true);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBulletedList("presentation", false, true);
            #endregion
            
            #region Generating Chart report with Filtering, Grouping, and Ordering
            //Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document", false, true);

            //Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet", false, true);

            //Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation", false, true);
            #endregion
            
            #region Generating Common List Report
            //Generate a Common List Report in document processing format
            //GenerateReport.GenerateCommonList("document", false, true);

            //Generate a Common List Report in spreadsheet format
            //GenerateReport.GenerateCommonList("spreadsheet", false, true);

            //Generate a Common List Report in presentation format
            //GenerateReport.GenerateCommonList("presentation", false, true);
            #endregion
            
            #region Generating Common Master-Detail Report
            //Generate a Common Master-Detail Report in document processing format
            //GenerateReport.GenerateCommonMasterDetail("document", false, true);

            //Generate a Common Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateCommonMasterDetail("spreadsheet", false, true);

            //Generate a Common Master-Detail Report in presentation format
            //GenerateReport.GenerateCommonMasterDetail("presentation", false, true);
            #endregion
            
            #region Generating In-Paragraph List Report
            //Generate a In-Paragraph List Report in document processing format
            //GenerateReport.GenerateInParagraphList("document", false, true);

            //Generate a In-Paragraph List Report in spreadsheet format
            //GenerateReport.GenerateInParagraphList("spreadsheet", false, true);

            //Generate a In-Paragraph List Report in presentation format
            //GenerateReport.GenerateInParagraphList("presentation", false, true);
            #endregion
            
            #region Generating In-Table with Alternate Content Report
            //Generate a In-Table List with Alternate Content Report in document processing format
            //GenerateReport.GenerateInTableListWithAlternateContent("document", false, true);

            //Generate a In-Table List with Alternate Content Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet", false, true);

            //Generate a In-Table List with Alternate Content Report in presentation format
            //GenerateReport.GenerateInTableListWithAlternateContent("presentation", false, true);
            #endregion
           
            #region Generating In-Table with Filtering, Grouping and Ordering Report
            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document", false, true);

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet", false, true);

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation", false, true);
            #endregion
            
            #region Generating In-Table List with Highlighted Rows Report
            //Generate a In-Table List with Highlighted Rows Report in document processing format
            //GenerateReport.GenerateInTableListWithHighlightedRows("document", false, true);

            //Generate a In-Table List with Highlighted Rows Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet", false, true);

            //Generate a In-Table List with Highlighted Rows Report in presentation format
            //GenerateReport.GenerateInTableListWithHighlightedRows("presentation", false, true);
            #endregion
            
            #region Generating In-Table List Report
            //Generate a In-Table List Report in document processing format
            //GenerateReport.GenerateInTableList("document", false, true);

            //Generate a In-Table List Report in spreadsheet format
            //GenerateReport.GenerateInTableList("spreadsheet", false, true);

            //Generate a In-Table List Report in presentation format
            //GenerateReport.GenerateInTableList("presentation", false, true);
            #endregion
            
            #region Generating In-Table Master-Detail Report
            //Generate a In-Table Master-Detail Report in document processing format
            //GenerateReport.GenerateInTableMasterDetail("document", false, true);

            //Generate a In-Table Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateInTableMasterDetail("spreadsheet", false, true);

            //Generate a In-Table Master-Detail Report in presentation format
            //GenerateReport.GenerateInTableMasterDetail("presentation", false, true);
            #endregion
            
            #region Generating Multicolored Number List Report
            //Generate a Multicolored Numbered List Report in document processing format
            //GenerateReport.GenerateMulticoloredNumberedList("document", false, true);

            //Generate a Multicolored Numbered List Report in spreadsheet format
            //GenerateReport.GenerateMulticoloredNumberedList("spreadsheet", false, true);

            //Generate a Multicolored Numbered List Report in presentation format
            //GenerateReport.GenerateMulticoloredNumberedList("presentation", false, true);
            #endregion
            
            #region Generating Numbered List Report
            //Generate a Numbered List Report in document processing format
            //GenerateReport.GenerateNumberedList("document", false, true);

            //Generate a Numbered List Report in spreadsheet format
            //GenerateReport.GenerateNumberedList("spreadsheet", false, true);

            //Generate a Numbered List Report in presentation format
            //GenerateReport.GenerateNumberedList("presentation", false, true);
            #endregion
            
            #region Generating Pie Chart Report
            //Generate a Pie Chart Report in document processing format
            //GenerateReport.GeneratePieChart("document", false, true);
            
            //Generate a Pie Chart Report in spreadsheet format
            //GenerateReport.GeneratePieChart("spreadsheet", false, true);

            //Generate a Pie Chart Report in presentation format
            //GenerateReport.GeneratePieChart("presentation", false, true);
            #endregion
            
            #region Generating Scatter Chart Report
            //Generate a Scatter Chart Report in document processing format
            //GenerateReport.GenerateScatterChart("document", false, true);

            //Generate a Scatter Chart Report in spreadsheet format
            //GenerateReport.GenerateScatterChart("spreadsheet", false, true);

            //Generate a Scatter Chart Report in presentation format
            //GenerateReport.GenerateScatterChart("presentation", false, true);
            #endregion
            
            #region Generating Single Row Report
            //Generate a Single Row Report in document processing format
            //GenerateReport.GenerateSingleRow("document", false, true);

            //Generate a Single Row Report in spreadsheet format
            //GenerateReport.GenerateSingleRow("spreadsheet", false, true);

            //Generate a Single Row Report in presentation format
            //GenerateReport.GenerateSingleRow("presentation", false, true);
            #endregion
            
           
        }
    }
}

