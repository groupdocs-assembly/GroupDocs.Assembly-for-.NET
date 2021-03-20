using System;
using GroupDocs.AssemblyExamples.BusinessLayer;

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
            CommonUtilities.ApplyLicense();

            //ExEnd:ApplyingLicense

            #region Change Target File Format

            //Change target file format using the file extension
            //GenerateReport.ChangeTargetFileFormat();

            //Change target file format using explicit specifying
            //GenerateReport.ChangeTargetFileFormatUsingExplicitSpecifying();
            #endregion

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

            //Generate a Nested Numbered List Report with restartNum in Documents
            //GenerateReport.GenerateNumberedListRestartNum_Documents();

            //Generate a Nested Numbered List Report with restartNum in Emails
            //GenerateReport.GenerateNumberedListRestartNum_Email();

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

            //Insert nested external output documents in word
            //GenerateReport.InsertNestedExternalDocumentsInWord();

            //Insert nested external output documents in email
            //GenerateReport.InsertNestedExternalDocumentsInEmail();

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
            // GenerateReport.RemoveSelectiveChartSeries();
            #endregion

            #region Dynamic Chart Axis Title 
            //Dynamic Chart Axis Title in Word Document 
            //GenerateReport.DynamicChartAxisTitle();
            //Dynamic Chart Axis Title in Presentation Document
            //GenerateReport.DynamicChartAxisTitlePPt();
            //Dynamic Chart Axis Title in Spreadsheet Document 
            //GenerateReport.DynamicChartAxisTitleSpreadSheet();
            //Dynamic Chart Axis Title in Presentation Document
            //GenerateReport.DynamicChartAxisTitleEmail();
            #endregion

            #region Dynamic Color 
            //GenerateReport.DynamicColor();
            // Sets colors of chart series dynamically based upon expressions wordprocessing document
            //GenerateReport.DynamicChartSeriesColor();
            // Sets colors of chart series dynamically based upon expressions spreadsheet document 
            //GenerateReport.DynamicChartSeriesColorSpreadsheet();
            // Sets colors of chart series dynamically based upon expressions presentation docuement
            //GenerateReport.DynamicChartSeriesColorPresentation();
            // Sets colors of chart series dynamically based upon expressions email docuement
            //GenerateReport.DynamicChartSeriesColorEmail();

            // Sets colors of chart series point color dynamically based upon expressions wordprocessing document
            //GenerateReport.DynamicChartSeriesPointColor();
            // Sets colors of chart series point color dynamically based upon expressions spreadsheet document 
            //GenerateReport.DynamicChartSeriesPointColorSpreadsheet();
            // Sets colors of chart series point color dynamically based upon expressions prosentation document
            //GenerateReport.DynamicChartSeriesPointColorPresentation();
            // Sets colors of chart series point color dynamically based upon expressions email document
            //GenerateReport.DynamicChartSeriesPointColorEmail();
            #endregion

            #region Working With Table Row Data Bands
            // Working With Table Row DataBands in Word Processing Document
            //GenerateReport.WorkingWithTableRowDataBandsWord();
            // Working With Table Row DataBands in SpreadSheet Document
            //GenerateReport.WorkingWithTableRowDataBandsSpreadSheet();
            // Working With Table Row DataBands in Presentation Document 
            //GenerateReport.WorkingWithTableRowDataBandsPresentation();
            // Working With Table Row DataBands in Email Format
            //GenerateReport.WorkingWithTableRowDataBandsEmail();
            #endregion

            #region Insert Hyperlinks Dynamically
            //Insert Hyperlink Dynamically in Word Document
            //GenerateReport.DynamicHyperlinkInsertionWord();
            //Insert Hyperlink Dynamically in Presentation Document
            //GenerateReport.DynamicHyperlinkInsertionPresentation();
            //Insert Hyperlink Dynamically in Spreadsheet Document
            //GenerateReport.DynamicHyperlinkInsertionSpreadsheet();
            //Insert Hyperlink Dynamically in Email Document
            //GenerateReport.DynamicHyperlinkInsertionEmail();

            #endregion

            #region Support removal of paragraphs becoming empty after template syntax tags are removed or replaced with empty values
            //Working with word processing document
            //GenerateReport.EmptyParagraphInWordProcessing();

            //Working with presentation document 
            //GenerateReport.EmptyParagraphInPresentation();

            //Working with email documents 
            //GenerateReport.EmptyParagraphInEmail();
            #endregion

            #region Merging table cells dynamically
            // Merging table cells dynamically in Word Processing
            //GenerateReport.TableCellsMergingInWordProcessing();

            // Merging table cells dynamically in Presentations
            //GenerateReport.TableCellsMergingInPresentations();

            // Merging table cells dynamically in Spreadsheets
            //GenerateReport.TableCellsMergingInSpreadsheets();

            // Merging table cells dynamically in Email
            // GenerateReport.TableCellsMergingInEmails();
            #endregion

            #region Using Markdown File Format
            //Saving an assembled Markdown document to a Word Processing format using file extension.
            //GenerateReport.SaveMdtoWord_UsingExtension("ReadMe.md");

            //Unordered lists demo for Markdown
            //GenerateReport.SaveMdtoWord_UsingExtension("List_demo.md");

            //Ordered lists demo for Markdown
            //GenerateReport.SaveMdtoWord_UsingExtension("Ordered list.md");

            //Saving an assembled Word Processing document or email to Markdown using file extension.
            //GenerateReport.SaveWordOrEmailtoMD_UsingExtension();

            //Saving an assembled Word Processing document or email to Markdown using file extension.
            //GenerateReport.SaveWordOrEmailtoMD_Explicit();
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

            //GenerateReport.UsingStringAsTemplate();

            // Demonstrate how to enable in-line syntax errors in the template without throw any exception
            //GenerateReport.DemoInLineSyntaxError();

            //Loading of template documents from HTML(with relative path) with resources
            //GenerateReport.LoadDocFromHTMLWithResource();

            // Loading of template documents from HTML with resources from an explicitly specified folder
            //GenerateReport.LoadDocFromHTMLWithResource_ExplicitFolder();

            //Saving of external resource files at relative path 
            //GenerateReport.SaveDocToHTMLWithResource();

            // Saving of external resource files in a specified folder at relative path while saving output to HTML
            //GenerateReport.SaveDocToHTMLWithResource_ExplicitFolder();

            //Working with JSON data sources
            //GenerateReport.SimpleJsonDS_Demo();

            //Working with XML data sources
            //GenerateReport.SimpleXMLDS_Demo();

            //Working with csv data sources
            //GenerateReport.SimpleCsvDS_Demo();

            //Insert Bookmarks Dynamically in Word Document
            //GenerateReport.DynamicBookmarkInsertionWord();

            //Insert Bookmarks Dynamically in Word Document
            // GenerateReport.DynamicBookmarkInsertionExcel();

            //Insert Image Dynamically in Word Document
            //GenerateReport.InsertImageDynamicallyInWord();

            // Set checkbox value dynamically in Word document
            //GenerateReport.SetCheckboxValueDynamicallyInWord(true);

            // Insert Document Dynamically in Word Document
            //GenerateReport.InsertDocumentDynamicallyInWord();

            // Set barcode resolution while saving.
            //GenerateReport.SetBarcodeResolution();

            #region Using Markdown File Format

            // Loading templates POT and OTP Presentation documents.
            //GenerateReport.SavePOTtoPPTX();
            //GenerateReport.SaveOTPtoPPTX();
            // Saving of assembled Presentation documents to POT and OTP formats.
            //GenerateReport.SavePPTXtoOTP();
            //GenerateReport.SavePPTXtoPOT();
            GenerateReport.SavePPTXtoPOTOTPAsStream();

            #endregion

            Console.WriteLine("Done...");
            Console.ReadKey();
        }
    }
}