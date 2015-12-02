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
            //GenerateReport.GenerateBubbleChart("document", false);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBubbleChart("spreadsheet", false);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBubbleChart("presentation", false);
            #endregion
            
            #region Generating Bulleted List Report
            //Generate a Bulleted List report in document processing format
            //GenerateReport.GenerateBulletedList("document", false);

            //Generate a Bulleted List report in spreadsheet format
            //GenerateReport.GenerateBulletedList("spreadsheet", false);

            //Generate a Bulleted List report in presentation format
            //GenerateReport.GenerateBulletedList("presentation", false);
            #endregion
            
            #region Generating Chart report with Filtering, Grouping, and Ordering
            //Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document", false);

            //Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet", false);

            //Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            //GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation", false);
            #endregion
            
            #region Generating Common List Report
            //Generate a Common List Report in document processing format
            //GenerateReport.GenerateCommonList("document", false);

            //Generate a Common List Report in spreadsheet format
            //GenerateReport.GenerateCommonList("spreadsheet", false);

            //Generate a Common List Report in presentation format
            //GenerateReport.GenerateCommonList("presentation", false);
            #endregion
            
            #region Generating Common Master-Detail Report
            //Generate a Common Master-Detail Report in document processing format
            //GenerateReport.GenerateCommonMasterDetail("document", false);

            //Generate a Common Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateCommonMasterDetail("spreadsheet", false);

            //Generate a Common Master-Detail Report in presentation format
            //GenerateReport.GenerateCommonMasterDetail("presentation", false);
            #endregion
            
            #region Generating In-Paragraph List Report
            //Generate a In-Paragraph List Report in document processing format
            //GenerateReport.GenerateInParagraphList("document", false);

            //Generate a In-Paragraph List Report in spreadsheet format
            //GenerateReport.GenerateInParagraphList("spreadsheet", false);

            //Generate a In-Paragraph List Report in presentation format
            //GenerateReport.GenerateInParagraphList("presentation", false);
            #endregion
            
            #region Generating In-Table with Alternate Content Report
            //Generate a In-Table List with Alternate Content Report in document processing format
            //GenerateReport.GenerateInTableListWithAlternateContent("document", false);

            //Generate a In-Table List with Alternate Content Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet", false);

            //Generate a In-Table List with Alternate Content Report in presentation format
            //GenerateReport.GenerateInTableListWithAlternateContent("presentation", false);
            #endregion
           
            #region Generating In-Table with Filtering, Grouping and Ordering Report
            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document", false);

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet", false);

            //Generate a In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            //GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation", false);
            #endregion
            
            #region Generating In-Table List with Highlighted Rows Report
            //Generate a In-Table List with Highlighted Rows Report in document processing format
            //GenerateReport.GenerateInTableListWithHighlightedRows("document", false);

            //Generate a In-Table List with Highlighted Rows Report in spreadsheet format
            //GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet", false);

            //Generate a In-Table List with Highlighted Rows Report in presentation format
            //GenerateReport.GenerateInTableListWithHighlightedRows("presentation", false);
            #endregion
            
            #region Generating In-Table List Report
            //Generate a In-Table List Report in document processing format
            //GenerateReport.GenerateInTableList("document", false);

            //Generate a In-Table List Report in spreadsheet format
            //GenerateReport.GenerateInTableList("spreadsheet", false);

            //Generate a In-Table List Report in presentation format
            //GenerateReport.GenerateInTableList("presentation", false);
            #endregion
            
            #region Generating In-Table Master-Detail Report
            //Generate a In-Table Master-Detail Report in document processing format
            //GenerateReport.GenerateInTableMasterDetail("document", false);

            //Generate a In-Table Master-Detail Report in spreadsheet format
            //GenerateReport.GenerateInTableMasterDetail("spreadsheet", false);

            //Generate a In-Table Master-Detail Report in presentation format
            //GenerateReport.GenerateInTableMasterDetail("presentation", false);
            #endregion
            
            #region Generating Multicolored Number List Report
            //Generate a Multicolored Numbered List Report in document processing format
            //GenerateReport.GenerateMulticoloredNumberedList("document", false);

            //Generate a Multicolored Numbered List Report in spreadsheet format
            //GenerateReport.GenerateMulticoloredNumberedList("spreadsheet", false);

            //Generate a Multicolored Numbered List Report in presentation format
            //GenerateReport.GenerateMulticoloredNumberedList("presentation", false);
            #endregion
            
            #region Generating Numbered List Report
            //Generate a Numbered List Report in document processing format
            //GenerateReport.GenerateNumberedList("document", false);

            //Generate a Numbered List Report in spreadsheet format
            //GenerateReport.GenerateNumberedList("spreadsheet", false);

            //Generate a Numbered List Report in presentation format
            //GenerateReport.GenerateNumberedList("presentation", false);
            #endregion
            
            #region Generating Pie Chart Report
            //Generate a Pie Chart Report in document processing format
            //GenerateReport.GeneratePieChart("document", false);
            
            //Generate a Pie Chart Report in spreadsheet format
            //GenerateReport.GeneratePieChart("spreadsheet", false);

            //Generate a Pie Chart Report in presentation format
            //GenerateReport.GeneratePieChart("presentation", false);
            #endregion
            
            #region Generating Scatter Chart Report
            //Generate a Scatter Chart Report in document processing format
            //GenerateReport.GenerateScatterChart("document", false);

            //Generate a Scatter Chart Report in spreadsheet format
            //GenerateReport.GenerateScatterChart("spreadsheet", false);

            //Generate a Scatter Chart Report in presentation format
            //GenerateReport.GenerateScatterChart("presentation", false);
            #endregion
            
            #region Generating Single Row Report
            //Generate a Single Row Report in document processing format
            //GenerateReport.GenerateSingleRow("document", false);

            //Generate a Single Row Report in spreadsheet format
            //GenerateReport.GenerateSingleRow("spreadsheet", false);

            //Generate a Single Row Report in presentation format
            //GenerateReport.GenerateSingleRow("presentation", false);
            #endregion
            
           
        }
    }
}

