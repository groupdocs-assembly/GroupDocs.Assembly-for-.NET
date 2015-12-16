using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.Assembly;
using GroupDocs.AssemblyExamples.BusinessLayer;
using GroupDocs.AssemblyExamples.ProjectEntities;

namespace GroupDocs.AssemblyExamples
{
    public static class GenerateReport
    {

        public static void GenerateBubbleChart(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBubbleChartFromDatabaseinOpenDocumentProcessingFormat
                        //setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bubble Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bubble Chart_DB Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
                            //Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateBubbleChartFromDataSetinOpenDocumentProcessingFormat
                        //setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bubble Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bubble Chart_DT Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
                            //Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateBubbleChartFromXMLinOpenDocumentProcessingFormat
                        //setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bubble Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bubble Chart_XML Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
                            //Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromXMLinOpenDocumentProcessingFormat
                        
                    }
                    else
                    {
                        //ExStart:GenerateBubbleChartinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bubble Chart.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bubble Chart Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBubbleChartFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart_DB Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateBubbleChartFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart_DT Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateBubbleChartFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart_XML Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateBubbleChartinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBubbleChartFromDatabaseinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Bubble Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Bubble Chart_DB Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateBubbleChartFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Bubble Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Bubble Chart_DT Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateBubbleChartFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Bubble Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Bubble Chart_XML Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateBubbleChartinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Bubble Chart.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Bubble Chart Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateBulletedList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBulletedListFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bulleted List_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateBulletedListFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bulleted List_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateBulletedListFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bulleted List_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bulleted List_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateBulletedListinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Bulleted List Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBulletedListFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bulleted List_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateBulletedListFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bulleted List_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateBulletedListFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bulleted List_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateBulletedListinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Bulleted List Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBulletedListFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Bulleted List_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateBulletedListFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Bulleted List_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateBulletedListFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Bulleted List_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Bulleted List_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateBulletedListinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Bulleted List Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateChartWithFilteringGroupingAndOrdering(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_DB Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_DT Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering_XML.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_XML Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx";
                        //Setting up destination open document report 
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
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DB Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DT Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_XML.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_XML Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx";
                        //Setting up destination open document report 
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
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DB Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DT Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_XML.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_XML Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx";
                        //Setting up destination open document report 
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
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateCommonList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonListFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common List_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateCommonListFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common List_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateCommonListFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common List_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateCommonListinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common List Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonListFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common List_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateCommonListFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common List_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateCommonListFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common List_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateCommonListinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common List Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonListFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Common List_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateCommonListFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Common List_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateCommonListFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Common List.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Common List_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateCommonListinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Common List.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Common List Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateCommonMasterDetail(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common Master-Detail_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common Master-Detail_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common Master-Detail_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common Master-Detail_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateCommonMasterDetailFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common Master-Detail_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common Master-Detail_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateCommonMasterDetailinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Common Master-Detail_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Common Master-Detail Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateCommonMasterDetailFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateCommonMasterDetailinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDatabaseinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Common Master-Detail_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Common Master-Detail_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateCommonMasterDetailFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Common Master-Detail_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateCommonMasterDetailinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Common Master-Detail_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Common Master-Detail Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateInParagraphList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInParagraphListFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Paragraph List_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInParagraphListFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Paragraph List_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInParagraphListFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Paragraph List_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Paragraph List_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateInParagraphListinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Paragraph List Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInParagraphListFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInParagraphListFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInParagraphListFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Paragraph List_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateInParagraphListinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInParagraphListFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Paragraph List_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInParagraphListFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Paragraph List_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInParagraphListFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Paragraph List_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Paragraph List_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateInParagraphListinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Paragraph List Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateInTableListWithAlternateContent(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Alternate Content_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content_DT_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Alternate Content_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Alternate Content_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithAlternateContentinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Alternate Content Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Alternate Content_DB_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Alternate Content_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Alternate Content_DT_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Alternate Content_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Alternate Content_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Alternate Content_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithAlternateContentinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Alternate Content Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Alternate Content_DB_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Alternate Content_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Alternate Content_DT_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Alternate Content_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Alternate Content_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Alternate Content_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithAlternateContentinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Alternate Content Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateInTableListWithFilteringGroupingAndOrdering(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinDocumentProcessingDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinDocumentProcessingDocument
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateInTableListWithHighlightedRows(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows_DT_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinDocumentProcessingDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromXMLinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinDocumentProcessingDocument
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows_DB_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Highlighted Rows_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinSpreadsheetDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows_DT_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Highlighted Rows_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinSpreadsheetDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromXMLinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Highlighted Rows_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinSpreadsheetDocument
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Highlighted Rows Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Highlighted Rows_DB_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Highlighted Rows_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Highlighted Rows_DT_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Highlighted Rows_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Highlighted Rows_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Highlighted Rows_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List with Highlighted Rows Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateInTableList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListFromDatabaseinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromDatabaseinDocumentProcessingDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListFromDataSetinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List_DT_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromDataSetinDocumentProcessingDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListFromXMLinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromXMLinDocumentProcessingDocument
                    }
                    else
                    {
                        //ExStart:GenerateInTableListinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table List Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListinDocumentProcessingDocument
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_DB_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_DT_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List_DB_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableListFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List_DT_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableListFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableListinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table List Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateInTableMasterDetail(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table Master-Detail_DB_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table Master-Detail_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table Master-Detail_DT_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table Master-Detail_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableMasterDetailFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table Master-Detail_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table Master-Detail_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableMasterDetailinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/In-Table Master-Detail_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/In-Table Master-Detail Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table Master-Detail_DB_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table Master-Detail_DT_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableMasterDetailFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table Master-Detail_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableMasterDetailinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/In-Table Master-Detail_DB_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table Master-Detail_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/In-Table Master-Detail_DT_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table Master-Detail_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateInTableMasterDetailFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/In-Table Master-Detail_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table Master-Detail_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateInTableMasterDetailinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/In-Table Master-Detail_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/In-Table Master-Detail Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateMulticoloredNumberedList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDatabaseinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Multicolored Numbered List_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDatabaseinDocumentProcessingDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDataSetinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Multicolored Numbered List_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDataSetinDocumentProcessingDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromXMLinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Multicolored Numbered List_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromXMLinDocumentProcessingDocument
                    }
                    else
                    {
                        //ExStart:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Multicolored Numbered List Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDatabaseinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDatabaseinSpreadsheetDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDataSetinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDataSetinSpreadsheetDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromXMLinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromXMLinSpreadsheetDocument
                    }
                    else
                    {
                        //ExStart:GenerateMulticoloredNumberedListinSpreadsheetDocument
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListinSpreadsheetDocument
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Multicolored Numbered List_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Multicolored Numbered List_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Multicolored Numbered List_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateMulticoloredNumberedListinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Multicolored Numbered List Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateNumberedList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateNumberedListFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Numbered List_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                       //ExStart:GenerateNumberedListFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Numbered List_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromDataSetinOpenDocumentProcessingFormat 
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateNumberedListFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Numbered List_XML_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Numbered List_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromXMLinOpenDocumentProcessingFormat 
                    }
                    else
                    {
                        //ExStart:GenerateNumberedListinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Numbered List Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateNumberedListFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Numbered List_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateNumberedListFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Numbered List_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateNumberedListFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_XML_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Numbered List_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateNumberedListinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        //Setting up destination open spreadsheet report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Numbered List Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateNumberedListFromDatabaseinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Numbered List_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateNumberedListFromDataSetinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Numbered List_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateNumberedListFromXMLinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Numbered List_XML_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Numbered List_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateNumberedListinOpenPresentationFormat
                        //Setting up source open presentation template
                        const String strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        //Setting up destination open presentation report 
                        const String strPresentationReport = "Presentation Reports/Numbered List Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GeneratePieChart(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GeneratePieChartFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Pie Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Pie Chart_DB Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GeneratePieChartFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Pie Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Pie Chart_DT Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GeneratePieChartFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Pie Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Pie Chart_XML Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GeneratePieChartinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Pie Chart.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Pie Chart Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GeneratePieChartFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Pie Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Pie Chart_DB Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GeneratePieChartFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Pie Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Pie Chart_DT Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GeneratePieChartFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Pie Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Pie Chart_XML Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GeneratePieChartinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Pie Chart.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Pie Chart Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GeneratePieChartFromDatabaseinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Pie Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Pie Chart_DB Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GeneratePieChartFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Pie Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Pie Chart_DT Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GeneratePieChartFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Pie Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Pie Chart_XML Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GeneratePieChartinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Pie Chart.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Pie Chart Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GeneratePieChartinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateScatterChart(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateScatterChartFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Scatter Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Scatter Chart_DB Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateScatterChartFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Scatter Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Scatter Chart_DT Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateScatterChartFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Scatter Chart_DB.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Scatter Chart_XML Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateScatterChartinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Scatter Chart.docx";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Scatter Chart Report.docx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateScatterChartFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Scatter Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Scatter Chart_DB Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateScatterChartFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Scatter Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Scatter Chart_DT Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateScatterChartFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Scatter Chart_DB.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Scatter Chart_XML Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateScatterChartinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Scatter Chart.xlsx";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Scatter Chart Report.xlsx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateScatterChartFromDatabaseinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Scatter Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Scatter Chart_DB Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateScatterChartFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Scatter Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Scatter Chart_DT Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateScatterChartFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Scatter Chart_DB.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Scatter Chart_XML Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateScatterChartinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Scatter Chart.pptx";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Scatter Chart Report.pptx";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateScatterChartinOpenPresentationFormat
                    }
                    break;
            }
        }
        public static void GenerateSingleRow(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateSingleRowFromDatabaseinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Single Row_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Single Row_DB Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDataDB(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromDatabaseinOpenDocumentProcessingFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateSingleRowFromDataSetinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Single Row_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Single Row_DT Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDT(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromDataSetinOpenDocumentProcessingFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateSingleRowFromXMLinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Single Row_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Single Row_XML Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerXML(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromXMLinOpenDocumentProcessingFormat
                    }
                    else
                    {
                        //ExStart:GenerateSingleRowinOpenDocumentProcessingFormat
                        //Setting up source open document template
                        const String strDocumentTemplate = "Word Templates/Single Row_OpenDocument.odt";
                        //Setting up destination open document report 
                        const String strDocumentReport = "Word Reports/Single Row Report.odt";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerData(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowinOpenDocumentProcessingFormat
                    }
                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateSingleRowFromDatabaseinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Single Row_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Single Row_DB Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerDataDB(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromDatabaseinOpenSpreadsheetFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateSingleRowFromDataSetinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Single Row_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Single Row_DT Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerDT(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromDataSetinOpenSpreadsheetFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateSingleRowFromXMLinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Single Row_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Single Row_XML Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerXML(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromXMLinOpenSpreadsheetFormat
                    }
                    else
                    {
                        //ExStart:GenerateSingleRowinOpenSpreadsheetFormat
                        //Setting up source open spreadsheet template
                        const String strSpreadsheetTemplate = "Spreadsheet Templates/Single Row_OpenDocument.ods";
                        //Setting up destination open document report 
                        const String strSpreadsheetReport = "Spreadsheet Reports/Single Row Report.ods";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomerData(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowinOpenSpreadsheetFormat
                    }
                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateSingleRowFromDatabaseinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Single Row_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Single Row_DB Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerDataDB(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromDatabaseinOpenPresentationFormat
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateSingleRowFromDataSetinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Single Row_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Single Row_DT Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerDT(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromDataSetinOpenPresentationFormat
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateSingleRowFromXMLinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Single Row_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Single Row_XML Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerXML(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowFromXMLinOpenPresentationFormat
                    }
                    else
                    {
                        //ExStart:GenerateSingleRowinOpenPresentationFormat
                        //Setting up source open spreadsheet template
                        const String strPresentationTemplate = "Presentation Templates/Single Row_OpenDocument.odp";
                        //Setting up destination open document report 
                        const String strPresentationReport = "Presentation Reports/Single Row Report.odp";
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomerData(), "customer");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateSingleRowinOpenPresentationFormat
                    }
                    break;
            }
        }
    }
}
