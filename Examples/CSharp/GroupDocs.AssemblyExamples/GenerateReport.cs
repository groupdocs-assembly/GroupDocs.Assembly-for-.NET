using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.Assembly;
using GroupDocs.AssemblyExamples.BusinessLayer;

namespace GroupDocs.AssemblyExamples
{
    public static class GenerateReport
    {
        
        public static void GenerateBubbleChart(string strDocumentFormat)
        {
            
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateBubbleChartinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Bubble Chart.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Bubble Chart Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Bubble Chart Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    { 
                        Console.WriteLine(ex.Message); 
                    }
                    //ExEnd:GenerateBubbleChartinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateBubbleChartinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Bubble Chart Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBubbleChartinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateBubbleChartinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Bubble Chart.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Bubble Chart Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Bubble Chart Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBubbleChartinPresentationFormat
                    break;
            }
        }

       
        public static void GenerateBulletedList(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateBulletedListinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Bulleted List.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Bulleted List Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Bulleted List Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBulletedListinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateBulletedListinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Bulleted List Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Bulleted List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBulletedListinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateBulletedListinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Bulleted List.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Bulleted List Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Bulleted List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBulletedListinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateChartWithFilteringGroupingAndOrdering(string strDocumentFormat)
        {
            
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateChartWithFilteringGroupingAndOrderinginDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateChartWithFilteringGroupingAndOrderinginSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateChartWithFilteringGroupingAndOrderinginPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginPresentationFormat
                    break;
            }
        }
        public static void GenerateCommonList(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateCommonListinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Common List.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Common List Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Common List Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonListinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateCommonListinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Common List.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Common List Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Common List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonListinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateCommonListinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Common List.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Common List Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Common List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonListinPresentationFormat
                    break;
            }
        }

    
        public static void GenerateCommonMasterDetail(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateCommonMasterDetailinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Common Master-Detail.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Common Master-Detail Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Common Master-Detail Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonMasterDetailinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateCommonMasterDetailinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Common Master-Detail.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Common List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonMasterDetailinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateCommonMasterDetailinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Common Master-Detail.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Common Master-Detail Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Common List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonMasterDetailinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateInParagraphList(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInParagraphListinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/In-Paragraph List.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/In-Paragraph List Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Paragraph List Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInParagraphListinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateInParagraphListinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Paragraph List.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Paragraph List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInParagraphListinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateInParagraphListinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/In-Paragraph List.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/In-Paragraph List Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Paragraph List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInParagraphListinPresentationFormat
                    break;
            }
        }
        public static void GenerateInTableListWithAlternateContent(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInTableListWithAlternateContentinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/In-Table List with Alternate Content Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Alternate Content Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithAlternateContentinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateInTableListWithAlternateContentinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Alternate Content.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Alternate Content Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Alternate Content Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithAlternateContentinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateInTableListWithAlternateContentinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/In-Table List with Alternate Content.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/In-Table List with Alternate Content Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Alternate Content Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithAlternateContentinPresentationFormat
                    break;
            }
        }
        public static void GenerateInTableListWithFilteringGroupingAndOrdering(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                    break;

                case "spreadsheet":
                    //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginPresentationFormat
                    break;
            }
        }

        
        public static void GenerateInTableListWithHighlightedRows(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                    break;

                case "spreadsheet":
                    //ExStart:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Highlighted Rows Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                    break;

                case "presentation":
                    //ExStart:GenerateInTableListWithHighlightedRowsinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/In-Table List with Highlighted Rows.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/In-Table List with Highlighted Rows Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithHighlightedRowsinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateInTableList(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInTableListinDocumentProcessingDocument
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/In-Table List.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/In-Table List Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListinDocumentProcessingDocument
                    break;

                case "spreadsheet":
                    //ExStart:GenerateInTableListinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateInTableListinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/In-Table List.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/In-Table List Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateInTableMasterDetail(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInTableMasterDetailinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/In-Table Master-Detail.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/In-Table Master-Detail Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table Master-Detail Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableMasterDetailinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateInTableMasterDetailinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table Master-Detail.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table Master-Detail Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableMasterDetailinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateInTableMasterDetailinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/In-Table Master-Detail.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/In-Table Master-Detail Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate In-Table Master-Detail Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableMasterDetailinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateMulticoloredNumberedList(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Multicolored Numbered List.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Multicolored Numbered List Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Multicolored Numbered List Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                    break;

                case "spreadsheet":
                    //ExStart:GenerateMulticoloredNumberedListinSpreadsheetDocument
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Multicolored Numbered List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateMulticoloredNumberedListinSpreadsheetDocument
                    break;

                case "presentation":
                    //ExStart:GenerateMulticoloredNumberedListinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Multicolored Numbered List Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Multicolored Numbered List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateMulticoloredNumberedListinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateNumberedList(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateNumberedListinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Numbered List.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Numbered List Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Numbered List Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateNumberedListinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateNumberedListinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Numbered List Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Numbered List Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateNumberedListinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateNumberedListinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Numbered List.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Numbered List Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Numbered List Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateNumberedListinPresentationFormat
                    break;
            }
        }

       
        public static void GeneratePieChart(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GeneratePieChartinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Pie Chart.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Pie Chart Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Pie Chart Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GeneratePieChartinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GeneratePieChartinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Pie Chart.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Pie Chart Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Pie Chart Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GeneratePieChartinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GeneratePieChartinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Pie Chart.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Pie Chart Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Pie Chart Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GeneratePieChartinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateScatterChart(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateScatterChartinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Scatter Chart.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Scatter Chart Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Scatter Chart Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateScatterChartinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateScatterChartinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Scatter Chart.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Scatter Chart Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Scatter Chart Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateScatterChartinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateScatterChartinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Scatter Chart.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Scatter Chart Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Scatter Chart Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateScatterChartinPresentationFormat
                    break;
            }
        }

        
        public static void GenerateSingleRow(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateSingleRowinDocumentProcessingFormat
                    //Setting up source document template
                    const String strDocumentTemplate = "Word Templates/Single Row.docx";
                    //Setting up destination document report 
                    const String strDocumentReport = "Word Reports/Single Row Report.docx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Single Row Report in document format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerData(), "customer");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateSingleRowinDocumentProcessingFormat
                    break;

                case "spreadsheet":
                    //ExStart:GenerateSingleRowinSpreadsheetFormat
                    //Setting up source spreadsheet template
                    const String strSpreadsheetTemplate = "Spreadsheet Templates/Single Row.xlsx";
                    //Setting up destination document report 
                    const String strSpreadsheetReport = "Spreadsheet Reports/Single Row Report.xlsx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Single Row Report in spreadsheet format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomerData(), "customer");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateSingleRowinSpreadsheetFormat
                    break;

                case "presentation":
                    //ExStart:GenerateSingleRowinPresentationFormat
                    //Setting up source spreadsheet template
                    const String strPresentationTemplate = "Presentation Templates/Single Row.pptx";
                    //Setting up destination document report 
                    const String strPresentationReport = "Presentation Reports/Single Row Report.pptx";
                    try
                    {
                        //Instantiate DocumentAssembler class
                        DocumentAssembler assembler = new DocumentAssembler();
                        //Call AssembleDocument to generate Single Row Report in presentation format
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomerData(), "customer");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateSingleRowinPresentationFormat
                    break;
            }
        }
    }
}
