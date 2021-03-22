using System;
using System.Text;
using GroupDocs.Assembly;
using GroupDocs.AssemblyExamples.BusinessLayer;
using System.Diagnostics;
using GroupDocs.Assembly.Data;
using System.IO;
using System.Xml;

namespace GroupDocs.AssemblyExamples
{
    public static class GenerateReport
    {
        /// <summary>
        /// Shows how to load document Table set using default options.
        /// Features is supported by version 17.01 or greater.
        /// </summary>
        /// <param name="dataSource">name of the data source file.</param>
        public static void LoadDocTableSet(string dataSource)
        {
            //ExStart:LoadDocTableSet
            // Load all document tables using default options.
            DocumentTableSet tableSet =
                new DocumentTableSet(CommonUtilities.GetDataSourceDocument("Word DataSource/" + dataSource));

            // Check loading.
            Debug.Assert(tableSet.Tables.Count == 3);
            Debug.Assert(tableSet.Tables[0].Name == "Table1");
            Debug.Assert(tableSet.Tables[1].Name == "Table2");
            Debug.Assert(tableSet.Tables[2].Name == "Table3");
            //ExEnd:LoadDocTableSet
        }

        /// <summary>
        /// Change target file format using explicit specifying.
        /// Features is supported by version 18.9 or greater.
        /// </summary>
        public static void ChangeTargetFileFormatUsingExplicitSpecifying()
        {
            // Setting up source open document template.
            const string strDocumentTemplate = "Word Templates/Bubble Chart.docx";
            // Setting up destination PDF report.
            const string strDocumentReport = "PDF Reports/Bubble Chart Report.pdf";
            try
            {
                // Instantiate DocumentAssembler class.
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Bubble Chart Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Change target file format using the file extension.
        /// Features is supported by version 18.9 or greater.
        /// </summary>
        public static void ChangeTargetFileFormat()
        {
            // Setting up source open document template.
            const string strDocumentTemplate = "Word Templates/Bubble Chart.docx";
            // Setting up destination PDF report.
            const string strDocumentReport = "PDF Reports/Bubble Chart Report.pdf";
            try
            {
                // Instantiate DocumentAssembler class.
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Bubble Chart Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Show how to Load document table set using custom options
        /// Features is supported by version 17.01 or greater
        /// </summary>
        /// <param name="dataSource">name of the data source file</param>
        public static void LoadDocTableSetWithCustomOptions(string dataSource)
        {
            //ExStart:LoadDocTableSetWithCustomOptions
            // Load document tables using custom options.
            DocumentTableSet tableSet = new DocumentTableSet(
                CommonUtilities.GetDataSourceDocument("Word DataSource/" + dataSource),
                new CustomDocumentTableLoadHandler());

            // Ensure that the second table is not loaded.
            Debug.Assert(tableSet.Tables.Count == 2);
            Debug.Assert(tableSet.Tables[0].Name == "Table1");
            Debug.Assert(tableSet.Tables[1].Name == "Table3");

            // Ensure that default options are used to load the first table (that is, default column names are used).
            Debug.Assert(tableSet.Tables[0].Columns.Count == 2);
            Debug.Assert(tableSet.Tables[0].Columns[0].Name == "Column1");
            Debug.Assert(tableSet.Tables[0].Columns[1].Name == "Column2");

            // Ensure that custom options are used to load the third table (that is, column names are extracted).
            Debug.Assert(tableSet.Tables[1].Columns.Count == 2);
            Debug.Assert(tableSet.Tables[1].Columns[0].Name == "Name");
            Debug.Assert(tableSet.Tables[1].Columns[1].Name == "Address");
            //ExEnd:LoadDocTableSetWithCustomOptions
        }

        /// <summary>
        /// Shows how to use document TableSet as DataSource.
        /// Features is supported by version 17.01 or greater.
        /// </summary>
        /// <param name="dataSource">Name of the data source file.</param>
        /// <param name="slideDoc">name of the template file.</param>
        public static void UseDocumentTableSetAsDataSource(string dataSource, string slideDoc)
        {
            //ExStart:UseDocumentTableSetAsDataSource
            const string outDocument = "Presentation Reports/Use Document Table Set As DataSource Output.pptx";
            string templateFile = CommonUtilities.GetSourceDocument("Presentation Templates/" + slideDoc);

            // Set table column names to be extracted from the document.
            DocumentTableSet tableSet = new DocumentTableSet(
                CommonUtilities.GetDataSourceDocument("Word DataSource/" + dataSource),
                new ColumnNameExtractingDocumentTableLoadHandler());

            // Set table names for conveniency.
            tableSet.Tables[0].Name = "Planets";
            tableSet.Tables[1].Name = "Persons";
            tableSet.Tables[2].Name = "Companies";

            // Pass DocumentTableSet as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(templateFile, CommonUtilities.SetDestinationDocument(outDocument),
                new DataSourceInfo(tableSet));
            //ExEnd:UseDocumentTableSetAsDataSource
        }

        /// <summary>
        /// Shows how to define document table relations.
        /// Feature is supported by version 17.01 or greater.
        /// </summary>
        /// <param name="relatedTables">name of the data source file.</param>
        /// <param name="docTableRelations">name of the template file.</param>
        public static void DefiningDocumentTableRelations(string relatedTables, string docTableRelations)
        {
            //ExStart:DefiningDocumentTableRelations
            const string outDocument = "Word Reports/document relations output.docx";
            string relatedTablesDataSource = CommonUtilities.GetDataSourceDocument("Excel DataSource/" + relatedTables);
            string templateFile = CommonUtilities.GetSourceDocument("Word Templates/" + docTableRelations);

            // Set table column names to be extracted from the document.
            DocumentTableSet tableSet = new DocumentTableSet(relatedTablesDataSource,
                new ColumnNameExtractingDocumentTableLoadHandler());

            // Define relations between tables.
            // NOTE: For Spreadsheet documents, table names are extracted from sheet names.
            tableSet.Relations.Add(
                tableSet.Tables["CLIENT"].Columns["ID"],
                tableSet.Tables["CONTRACT"].Columns["CLIENT_ID"]);

            tableSet.Relations.Add(
                tableSet.Tables["MANAGER"].Columns["ID"],
                tableSet.Tables["CONTRACT"].Columns["MANAGER_ID"]);

            // Pass DocumentTableSet as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(templateFile, CommonUtilities.SetDestinationDocument(outDocument),
                new DataSourceInfo(tableSet));
            //ExEnd:DefiningDocumentTableRelations
        }

        /// <summary>
        /// Shows how to change document table column type.
        /// Feature is supported by version 17.01 or greater.
        /// </summary>
        /// <param name="document">Document path.</param>
        public static void ChangingDocumentTableColumnType(string document)
        {
            //ExStart:ChangingDocumentTableColumnType
            const string dataSrcDocument = "Word DataSource/Managers Data.docx";
            const string outDocument = "Presentation Reports/Out.pptx";

            // Set table column names to be extracted from the document.
            DocumentTableOptions options = new DocumentTableOptions {FirstRowContainsColumnNames = true};

            DocumentTable table = new DocumentTable(CommonUtilities.GetDataSourceDocument(dataSrcDocument), 1, options);

            // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
            Debug.Assert(table.Columns["Total_Contract_Price"].Type == typeof(string));

            // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
            // such as summing in templates.
            table.Columns["Total_Contract_Price"].Type = typeof(double);

            // Pass DocumentTable as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(document),
                CommonUtilities.SetDestinationDocument(outDocument), new DataSourceInfo(table, "Managers"));
            //ExEnd:ChangingDocumentTableColumnType
        }

        public static void GenerateBubbleChart(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBubbleChartFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Bubble Chart_DB.docx";
                        const string strDocumentReport = "Word Reports/Bubble Chart_DB Report.docx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strDocumentTemplate = "Word Templates/Bubble Chart_DB.docx";
                        const string strDocumentReport = "Word Reports/Bubble Chart_DT Report.docx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strDocumentTemplate = "Word Templates/Bubble Chart_DB.docx";
                        const string strDocumentReport = "Word Reports/Bubble Chart_XML Report.docx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            //Call AssembleDocument to generate Bubble Chart Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromXMLinOpenDocumentProcessingFormat

                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateBubbleChartFromJsoninWord
                        const string strDocumentTemplate = "Word Templates/Bubble Chart.docx";
                        const string strDocumentReport = "Word Reports/Bubble Chart_Json Report.docx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                        //ExEnd:GenerateBubbleChartFromJsoninWord
                    }
                    else
                    {
                        //ExStart:GenerateBubbleChartinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Bubble Chart.docx";
                        const string strDocumentReport = "Word Reports/Bubble Chart Report.docx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart_DB.xlsx";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart_DB Report.xlsx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart_DB.xlsx";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart_DT Report.xlsx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart_DB.xlsx";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart_XML Report.xlsx";

                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateBubbleChartFromJsoninExcel
                        const string strDocumentTemplate = "Spreadsheet Templates/Bubble Chart.xlsx";
                        const string strDocumentReport = "Spreadsheet Reports/Bubble Chart_Json Report.xlsx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromJsoninExcel
                    }
                    else
                    {
                        //ExStart:GenerateBubbleChartinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart.xlsx";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart Report.xlsx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strPresentationTemplate = "Presentation Templates/Bubble Chart_DB.pptx";
                        const string strPresentationReport = "Presentation Reports/Bubble Chart_DB Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strPresentationTemplate = "Presentation Templates/Bubble Chart_DB.pptx";
                        const string strPresentationReport = "Presentation Reports/Bubble Chart_DT Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strPresentationTemplate = "Presentation Templates/Bubble Chart_DB.pptx";
                        const string strPresentationReport = "Presentation Reports/Bubble Chart_XML Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateBubbleChartFromJsoninPresentation
                        const string strDocumentTemplate = "Presentation Templates/Bubble Chart.pptx";
                        const string strDocumentReport = "Presentation Reports/Bubble Chart_Json Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartFromJsoninPresentation
                    }
                    else
                    {
                        //ExStart:GenerateBubbleChartinOpenPresentationFormat
                        const string strPresentationTemplate = "Presentation Templates/Bubble Chart.pptx";
                        const string strPresentationReport = "Presentation Reports/Bubble Chart Report.pptx";
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bubble Chart Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBubbleChartinOpenPresentationFormat
                    }

                    break;
                case "email":
                {
                    //ExStart:GenerateBubbleChartinEmailFormat
                    const string strEmailTemplate = "Email Templates/Bubble Chart.msg";
                    const string strEmailReport = "Email Reports/Bubble Chart Report.msg";
                    
                    try
                    {
                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Bubble Chart Report in open presentation format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate), CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBubbleChartinEmailFormat
                }
                    break;
            }
        }

        /// <summary>
        /// This method inertes nested external doucments in email document.
        /// </summary>
        public static void InsertNestedExternalDocumentsInEmail()
        {
            const string strDocumentTemplate = "Email Templates/Nested External Document.msg";
            const string strDocumentReport = "Email Reports/Nested External Document.msg";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate  Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// This method inertes nested external doucments in word document.
        /// </summary>
        public static void InsertNestedExternalDocumentsInWord()
        {
            const string strDocumentTemplate = "Word Templates/Nested External Document.docx";
            const string strDocumentReport = "Word Reports/Nested External Document.docx";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate  Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static void EmptyParagraphInEmail()
        {
            const string strDocumentTemplate = "Email Templates/Empty Paragraph.msg";
            const string strDocumentReport = "Email Reports/Empty Paragraph.msg";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler
                {
                    // Apply Remove Empty Paragraph option.
                    Options = DocumentAssemblyOptions.RemoveEmptyParagraphs
                };
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("dummy", "dummy"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void EmptyParagraphInPresentation()
        {
            const string strDocumentTemplate = "Presentation Templates/Empty Paragraph.pptx";
            const string strDocumentReport = "Presentation Reports/Empty Paragraph.pptx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler
                {
                    // Apply Remove Empty Paragraph option.
                    Options = DocumentAssemblyOptions.RemoveEmptyParagraphs
                };
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("dummy", "dummy"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Supports removal of paragraphs becoming empty after template syntax tags are removed or replaced with empty values for Word Processing documents.
        /// </summary>
        public static void EmptyParagraphInWordProcessing()
        {
            const string strDocumentTemplate = "Word Templates/Empty Paragraph.docx";
            const string strDocumentReport = "Word Reports/Empty Paragraph.docx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler
                {
                    // Apply Remove Empty Paragraph option.
                    Options = DocumentAssemblyOptions.RemoveEmptyParagraphs
                };
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("dummy", "dummy"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Insert Hyperlink Dynamically in Email Document.
        /// Feature is supported by version 18.7 or greater.
        /// </summary>
        public static void DynamicHyperlinkInsertionEmail()
        {
            const string strDocumentTemplate = "Email Templates/Dynamic Hyperlink.msg";
            const string strDocumentReport = "Email Reports/Dynamic Hyperlink.msg";
            const string uriExpression = "https://www.groupdocs.com/";
            const string displayTextExpression = "GroupDocs";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to assemble document. 
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(uriExpression, "uriExpression"),
                    new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Insert Hyperlink Dynamically in Spreadsheet Document.
        /// Feature is supported by version 18.7 or greater.
        /// </summary>
        public static void DynamicHyperlinkInsertionSpreadsheet()
        {
            const string strDocumentTemplate = "Spreadsheet Templates/Dynamic Hyperlink.xlsx";
            const string strDocumentReport = "Spreadsheet Reports/Dynamic Hyperlink.xlsx";
            const string uriExpression = "https://www.groupdocs.com/";
            const string displayTextExpression = "GroupDocs";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(uriExpression, "uriExpression"),
                    new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Insert Hyperlink Dynamically in Presentaion Document.
        /// Feature is supported by version 18.7 or greater.
        /// </summary>
        public static void DynamicHyperlinkInsertionPresentation()
        {
            const string strDocumentTemplate = "Presentation Templates/Dynamic Hyperlink.pptx";
            const string strDocumentReport = "Presentation Reports/Dynamic Hyperlink.pptx";
            const string uriExpression = "https://www.groupdocs.com/";
            const string displayTextExpression = "GroupDocs";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(uriExpression, "uriExpression"),
                    new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Insert Hyperlink Dynamically in Word Document.
        /// Feature is supported by version 18.7 or greater.
        /// </summary>
        public static void DynamicHyperlinkInsertionWord()
        {
            const string strDocumentTemplate = "Word Templates/Dynamic Hyperlink.docx";
            const string strDocumentReport = "Word Reports/Dynamic Hyperlink.docx";
            const string uriExpression = "https://www.groupdocs.com/";
            const string displayTextExpression = "GroupDocs";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(uriExpression, "uriExpression"),
                    new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Insert Bookmarks Dynamically in Word Document.
        /// Feature is supported by version 20.1 or greater.
        /// </summary>
        public static void DynamicBookmarkInsertionWord()
        {
            const string strDocumentTemplate = "Word Templates/Dynamic bookmarks.docx";
            const string strDocumentReport = "Word Reports/Dynamic bookmarks.docx";
            const string bookmark_expression = "gd_bookmark";
            const string displayTextExpression = "GroupDocs";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(bookmark_expression, "bookmark_expression"),
                    new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Insert Bookmarks Dynamically in Excel Document.
        /// Feature is supported by version 20.1 or greater.
        /// </summary>
        public static void DynamicBookmarkInsertionExcel()
        {
            const string strDocumentTemplate = "Spreadsheet Templates/Dynamic Cell Range.xlsx";
            const string strDocumentReport = "Spreadsheet Reports/Dynamic Cell Range.xlsx";
            const string bookmark_expression = "gd_bookmark";
            const string displayTextExpression = "GroupDocs";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to assemble document.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(bookmark_expression, "bookmark_expression"),
                    new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series point color dynamically based upon expressions.
        /// Feature is supported by version 18.6 or greater.
        /// </summary>
        public static void DynamicChartSeriesPointColorEmail()
        {
            const string strDocumentTemplate = "Email Templates/Dynamic Chart Point Series Color.msg";
            const string strDocumentReport = "Email Reports/Dynamic Chart Point Series Color.msg";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Pie Chart report in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series color dynamically based upon expressions.
        /// Feature is supported by version 18.6 or greater.
        /// </summary>
        public static void DynamicChartSeriesColorEmail()
        {
            const string strDocumentTemplate = "Email Templates/Dynamic Chart Series Color.msg";
            const string dataSrcDocument = "Word DataSource/Managers Data.docx";
            const string strDocumentReport = "Email Reports/Dynamic Chart Series Color.msg";
            string color = "red";
            
            try
            {
                // Set table column names to be extracted from the document.
                DocumentTableOptions options = new DocumentTableOptions();
                options.FirstRowContainsColumnNames = true;

                DocumentTable table = new DocumentTable(CommonUtilities.GetDataSourceDocument(dataSrcDocument), 1, options);

                // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
                Debug.Assert(table.Columns["Total_Contract_Price"].Type == typeof(string));
                // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
                // such as summing in templates.
                table.Columns["Total_Contract_Price"].Type = typeof(double);

                // Pass DocumentTable as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(table, "managers"),
                    new DataSourceInfo(color, "color"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series point color dynamically based upon expressions.
        /// Feature is supported by version 18.5 or greater.
        /// </summary>
        public static void DynamicChartSeriesPointColorPresentation()
        {
            const string strDocumentTemplate = "Presentation Templates/Dynamic Chart Point Series Color.pptx";
            const string strDocumentReport = "Presentation Reports/Dynamic Chart Point Series Color.pptx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Pie Chart report in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series point color dynamically based upon expressions.
        /// Feature is supported by version 18.5 or greater.
        /// </summary>
        public static void DynamicChartSeriesPointColorSpreadsheet()
        {
            const string strDocumentTemplate = "Spreadsheet Templates/Dynamic Chart Point Series Color.xlsx";
            const string strDocumentReport = "Spreadsheet Reports/Dynamic Chart Point Series Color.xlsx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Pie Chart report in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series point color dynamically based upon expressions.
        /// Feature is supported by version 18.5 or greater.
        /// </summary>
        public static void DynamicChartSeriesPointColor()
        {
            const string strDocumentTemplate = "Word Templates/Dynamic Chart Point Series Color.docx";
            const string strDocumentReport = "Word Reports/Dynamic Chart Point Series Color.docx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Pie Chart report in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series dynamically based upon expressions.
        /// Feature is supported by version 18.5 or greater.
        /// </summary>
        public static void DynamicChartSeriesColorPresentation()
        {
            try
            {
                const string strDocumentTemplate = "Presentation Templates/Dynamic Chart Series Color.pptx";
                const string dataSrcDocument = "Presentation DataSource/Managers Data.pptx";
                const string strDocumentReport = "Presentation Reports/Dynamic Chart Series Color.pptx";
                string color = "red";

                // Set table column names to be extracted from the document.
                DocumentTableOptions options = new DocumentTableOptions
                {
                    FirstRowContainsColumnNames = true
                };

                DocumentTable table = new DocumentTable(CommonUtilities.GetDataSourceDocument(dataSrcDocument), 1, options);

                // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
                Debug.Assert(table.Columns["Total_Contract_Price"].Type == typeof(string));
                // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
                // such as summing in templates.
                table.Columns["Total_Contract_Price"].Type = typeof(double);

                // Pass DocumentTable as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(table, "managers"),
                    new DataSourceInfo(color, "color"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series dynamically based upon expressions.
        /// Feature is supported by version 18.5 or greater.
        /// </summary>
        public static void DynamicChartSeriesColorSpreadsheet()
        {
            try
            {
                const string strDocumentTemplate = "Spreadsheet Templates/Dynamic Chart Series Color.xlsx";
                const string dataSrcDocument = "Word DataSource/Managers Data.docx";
                const string strDocumentReport = "Spreadsheet Reports/Dynamic Chart Series Color.xlsx";
                string color = "red";

                // Set table column names to be extracted from the document.
                DocumentTableOptions options = new DocumentTableOptions();
                options.FirstRowContainsColumnNames = true;

                DocumentTable table = new DocumentTable(CommonUtilities.GetDataSourceDocument(dataSrcDocument), 1, options);

                // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
                Debug.Assert(table.Columns["Total_Contract_Price"].Type == typeof(string));
                // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
                // such as summing in templates.
                table.Columns["Total_Contract_Price"].Type = typeof(double);

                // Pass DocumentTable as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(table, "managers"),
                    new DataSourceInfo(color, "color"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Sets colors of chart series dynamically based upon expressions.
        /// Feature is supported by version 18.5 or greater.
        /// </summary>
        public static void DynamicChartSeriesColor()
        {
            try
            {
                const string strDocumentTemplate = "Word Templates/Dynamic Chart Series Color.docx";
                const string dataSrcDocument = "Word DataSource/Managers Data.docx";
                const string strDocumentReport = "Word Reports/Dynamic Chart Series Color.docx";
                string color = "red";

                // Set table column names to be extracted from the document.
                DocumentTableOptions options = new DocumentTableOptions();
                options.FirstRowContainsColumnNames = true;

                DocumentTable table = new DocumentTable(CommonUtilities.GetDataSourceDocument(dataSrcDocument), 1, options);

                // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
                Debug.Assert(table.Columns["Total_Contract_Price"].Type == typeof(string));
                // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
                // such as summing in templates.
                table.Columns["Total_Contract_Price"].Type = typeof(double);

                // Pass DocumentTable as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), 
                    new DataSourceInfo(table, "managers"),
                    new DataSourceInfo(color, "color"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void UsingStringAsTemplate()
        {
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                string sourceString = @"<<[yourValue]>>";
                byte[] sourceBytes = Encoding.UTF8.GetBytes(sourceString);
                byte[] targetBytes;

                using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
                {
                    using (MemoryStream targetStream = new MemoryStream())
                    {
                        assembler.AssembleDocument(sourceStream, targetStream,
                            new DataSourceInfo("Hello, World!", "yourValue"));
                        targetBytes = targetStream.ToArray();
                    }
                }

                string targetString = Encoding.UTF8.GetString(targetBytes);
                Console.WriteLine(targetString);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Working With Table Row DataBands in Email Message.
        /// Feature is supported by version 18.2 or greater.
        /// </summary>
        public static void WorkingWithTableRowDataBandsEmail()
        {
            //ExStart:WorkingWithTableRowDataBandsEmail
            string strDocumentTemplate = "Email Templates/Working With Table Row Data Bands.msg";
            string strDocumentReport = "Email Reports/Working With Table Row Data Bands.msg";
            
            try
            {
                // Assemble a document using the external document table as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.ExcelData(), "ds"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:WorkingWithTableRowDataBandsEmail
        }

        /// <summary>
        /// Working With Table Row DataBands in Presentation Documents.
        /// Feature is supported by version 18.2 or greater.
        /// </summary>
        public static void WorkingWithTableRowDataBandsPresentation()
        {
            //ExStart:WorkingWithTableRowDataBandsPresentation
            string strDocumentTemplate = "Presentation Templates/Working With Table Row Data Bands.pptx";
            string strDocumentReport = "Presentation Reports/Working With Table Row Data Bands.pptx";
            
            try
            {
                // Assemble a document using the external document table as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.ExcelData(), "ds"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:WorkingWithTableRowDataBandsPresentation
        }

        /// <summary>
        /// Working With Table Row DataBands in Spread Sheet Document.
        /// Feature is supported by version 18.2 or greater.
        /// </summary>
        public static void WorkingWithTableRowDataBandsSpreadSheet()
        {
            //ExStart:WorkingWithTableRowDataBandsSpreadSheet
            string strDocumentTemplate = "Spreadsheet Templates/Working With Table Row Data Bands.xlsx";
            string strDocumentReport = "Spreadsheet Reports/Working With Table Row Data Bands.xlsx";
            
            try
            {
                // Assemble a document using the external document table as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.ExcelData(), "ds"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:WorkingWithTableRowDataBandsSpreadSheet
        }

        /// <summary>
        /// Working With Table Row DataBands in Words Processing Document.
        /// Feature is supported by version 18.2 or greater.
        /// </summary>
        public static void WorkingWithTableRowDataBandsWord()
        {
            //ExStart:WorkingWithTableRowDataBandsWord
            string strDocumentTemplate = "Word Templates/Working With Table Row Data Bands.docx";
            string strDocumentReport = "Word Reports/Working With Table Row Data Bands.docx";
            
            try
            {
                // Assemble a document using the external document table as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.ExcelData(), "ds"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:WorkingWithTableRowDataBandsWord
        }

        public static void DynamicChartAxisTitleEmail()
        {
            //ExStart:DynamicChartAxisTitleEmail
            const string strEmailTemplate =
                "Email Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.msg";
            const string strEmailReport =
                "Email Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.msg";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                string title = "Total Order Quantity by Quarters";

                var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData(), title);
                var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders", "title");

                // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in email format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate), CommonUtilities.SetDestinationDocument(strEmailReport),
                    new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                    new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                    new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                    new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                    new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject),
                    new DataSourceInfo(dataSources.Title, dataSourcesNames.Title)
                );

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:DynamicChartAxisTitleEmail
        }

        public static void DynamicChartAxisTitleSpreadSheet()
        {
            //ExStart:DynamicChartAxisTitleSpreadSheet
            const string strDocumentTemplate =
                "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.xlsx";
            const string strDocumentReport =
                "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.xlsx";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                string title = "Total Order Quantity by Quarters";
                // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), 
                    new DataSourceInfo(title, "title"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:DynamicChartAxisTitleSpreadSheet
        }

        public static void DynamicChartAxisTitlePPt()
        {
            //ExStart:DynamicChartAxisTitlePPt
            const string strDocumentTemplate =
                "Presentation Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.pptx";
            const string strDocumentReport =
                "Presentation Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.pptx";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                string title = "Total Order Quantity by Quarters";
                // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), 
                    new DataSourceInfo(title, "title"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:DynamicChartAxisTitlePPt
        }

        public static void DynamicColor()
        {
            //ExStart:DynamicColor
            const string strDocumentTemplate =
                "Word Templates/In-Table List with Running (Progressive) Total_BackgroundColor.docx";
            const string strDocumentReport =
                "Word Reports/In-Table List with Running (Progressive) Total_BackgroundColor.docx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                string color = "red";
                // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), 
                    new DataSourceInfo(color, "color"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:DynamicColor
        }

        public static void DynamicChartAxisTitle()
        {
            //ExStart:DynamicChartAxisTitle
            const string strDocumentTemplate =
                "Word Templates/Chart with Filtering, Grouping, and Ordering_dynamic_title.docx";
            const string strDocumentReport =
                "Word Reports/Chart with Filtering, Grouping, and Ordering_dynamic_title.docx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                string title = "Total Order Quantity by Quarters";
                // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), 
                    new DataSourceInfo(title, "title"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:DynamicChartAxisTitle
        }

        public static void RemoveSelectiveChartSeries()
        {
            //ExStart:RemoveSelectiveChartSeries
            const string strDocumentTemplate =
                "Word Templates/Chart with Filtering, Grouping, and Ordering_RemoveIf.docx";
            const string strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_RemoveIf.docx";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                int mode = 2;
                // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), 
                    new DataSourceInfo(mode, "mode"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:RemoveSelectiveChartSeries
        }

        public static void GenerateBulletedList(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateBulletedListFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Bulleted List_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Bulleted List_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strDocumentTemplate = "Word Templates/Bulleted List_XML_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Bulleted List_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateBulletedListFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Bulleted List_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted list report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateBulletedListinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Bulleted List Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bulleted List_DB Report.ods";
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bulleted List_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Bulleted List_XML_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bulleted List_XML Report.ods";
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateBulletedListFromJsoninOpenExcel
                        const string strDocumentTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/Bulleted List_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler
                                assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate bulleted list in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromJsoninOpenExcel
                    }
                    else
                    {
                        //ExStart:GenerateBulletedListinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods"; 
                        const string strSpreadsheetReport = "Spreadsheet Reports/Bulleted List Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
                        const string strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Bulleted List_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Bulleted List_DT Report.odp";
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Bulleted List_XML_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Bulleted List_XML Report.odp";
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateBulletedListFromJsoninOpenPresentation
                        const string strDocumentTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/Bulleted List_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate bulleted list in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateBulletedListinOpenPresentationFormat
                        const string strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp"; 
                        const string strPresentationReport = "Presentation Reports/Bulleted List Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Bulleted List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateBulletedListinOpenPresentationFormat
                    }

                    break;

                case "html":
                {
                    //ExStart:GenerateBulletedListinHTMLFormat
                    const string strHtmlTemplate = "HTML Templates/Bulleted List.html";
                    const string strHtmlReport = "HTML Reports/Bulleted List Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // This is needed solely for images in HTML documents.
                        assembler.KnownTypes.Add(typeof(CommonUtilities.FileUtil));
                        // Call AssembleDocument to generate Bulleted List Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBulletedListinHTMLFormat
                }
                    break;
                case "text":
                {
                    //ExStart:GenerateBulletedListinTextFormat
                    const string strTextTemplate = "Text Templates/Bulleted List.txt";
                    const string strTextReport = "Text Reports/Bulleted List Report.txt";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Bulleted List Report in text format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strTextTemplate),
                            CommonUtilities.SetDestinationDocument(strTextReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBulletedListinTextFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateBulletedListinEmailFormat
                    const string strEmailTemplate = "Email Templates/Bulleted List.msg";
                    const string strEmailReport = "Email Reports/Bulleted List Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources =
                            DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetProductsData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

                        // Call AssembleDocument to generate Bulleted List Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate), CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateBulletedListinEmailFormat
                }
                    break;

            }
        }

        public static void GenerateChartWithFilteringGroupingAndOrdering(string strDocumentFormat, bool isDatabase,
            bool isDataSet, bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate =
                            "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx";
                        const string strDocumentReport =
                            "Word Reports/Chart with Filtering, Grouping, and Ordering_DB Report.docx";
                        try
                        { 
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strDocumentTemplate =
                            "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx";
                        const string strDocumentReport =
                            "Word Reports/Chart with Filtering, Grouping, and Ordering_DT Report.docx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strDocumentTemplate =
                            "Word Templates/Chart with Filtering, Grouping, and Ordering_XML.docx";
                        const string strDocumentReport =
                            "Word Reports/Chart with Filtering, Grouping, and Ordering_XML Report.docx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromJsoninWord
                        const string strDocumentTemplate =
                            "Word Templates/Chart with Filtering, Grouping, and Ordering.docx";
                        const string strDocumentReport =
                            "Word Reports/Chart with Filtering, Grouping, and Ordering_Json Report.docx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromJsoninWord
                    }
                    else
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenDocumentProcessingFormat
                        const string strDocumentTemplate =
                            "Word Templates/Chart with Filtering, Grouping, and Ordering.docx";
                        const string strDocumentReport =
                            "Word Reports/Chart with Filtering, Grouping, and Ordering Report.docx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DB Report.xlsx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DT Report.xlsx";
                        
                        try
                        {
                            //Instantiate DocumentAssembler class
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_XML.xlsx";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_XML Report.xlsx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromJsoninSpreadsheet
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx";
                        const string strDocumentReport =
                            "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_Json Report.xlsx";
                        
                        try
                        {
                            DocumentAssembler
                                assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromJsoninSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering Report.xlsx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx";
                        const string strPresentationReport =
                            "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DB Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx";
                        const string strPresentationReport =
                            "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DT Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Chart with Filtering, Grouping, and Ordering_XML.pptx";
                        const string strPresentationReport =
                            "Presentation Reports/Chart with Filtering, Grouping, and Ordering_XML Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderingFromJsoninPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx";
                        const string strDocumentReport =
                            "Presentation Reports/Chart with Filtering, Grouping, and Ordering_Json Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromJsoninPresentation
                    }
                    else
                    {
                        //ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx";
                        const string strPresentationReport =
                            "Presentation Reports/Chart with Filtering, Grouping, and Ordering Report.pptx";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenPresentationFormat
                    }

                    break;
                case "email":
                {
                    //ExStart:GenerateChartWithFilteringGroupingAndOrderinginEmailFormat
                    const string strEmailTemplate = "Email Templates/Chart with Filtering, Grouping, and Ordering.msg";
                    const string strEmailReport =
                        "Email Reports/Chart with Filtering, Grouping, and Ordering Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

                        // Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateChartWithFilteringGroupingAndOrderinginEmailFormat
                }
                    break;

            }
        }

        public static void GenerateCommonList(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonListFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common List_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common List_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common List_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateCommonListReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common List_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateCommonListinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common List Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common List_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common List_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common List_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateCommonListReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        const string strDocumentReport = "Spreadsheet Reports/Common List_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateCommonListinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common List Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common List_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common List_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common List_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateCommonListReportFromJsoninOpenPresentation
                        const string strDocumentTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/Common List_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List report in open presentation format.s
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateCommonListinOpenPresentationFormat
                        const string strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common List Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonListinOpenPresentationFormat
                    }

                    break;

                case "html":
                {
                    //ExStart:GenerateCommonListinHtmlFormat
                    const string strDocumentTemplate = "HTML Templates/Common List.html";
                    const string strDocumentReport = "HTML Reports/Common List Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Common List Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonListinHtmlFormat
                }
                    break;
                case "text":
                {
                    //ExStart:GenerateCommonListinTextFormat
                    const string strTxtDocumentTemplate = "Text Templates/Common List.txt";
                    const string strTxtDocumentReport = "Text Reports/Common List Report.txt";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Common List Report in text document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strTxtDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strTxtDocumentReport),
                            new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonListinTextFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateCommonListinEmailFormat
                    const string strEmailDocumentTemplate = "Email Templates/Common List.msg";
                    const string strEmailDocumentReport = "Email Reports/Common List Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources =
                            DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.PopulateData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");
                        // Call AssembleDocument to generate Common List Report in email document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailDocumentReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonListinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateCommonMasterDetail(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateCommonMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Common Master-Detail_DB_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common Master-Detail_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strDocumentTemplate = "Word Templates/Common Master-Detail_DB_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common Master-Detail_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strDocumentTemplate = "Word Templates/Common Master-Detail_DB_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common Master-Detail_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateCommonMasterDetailReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/Common Master-Detail_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/Common Master-Detail_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common master-detail report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateCommonMasterDetailinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Common Master-Detail_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Common Master-Detail Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateCommonMasterDetailReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/Common Master-Detail_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common master-detail report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateCommonMasterDetailinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common Master-Detail_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common Master-Detail_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common Master-Detail_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateCommonMasterDetailReportFromJsoninOpenPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/Common Master-Detail_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/Common Master-Detail_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common master-detail report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateCommonMasterDetailinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/Common Master-Detail_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Common Master-Detail Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Common List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateCommonMasterDetailinOpenPresentationFormat
                    }

                    break;

                case "html":
                {
                    //ExStart:GenerateCommonMasterDetailinHtmlFormat
                    const string strDocumentTemplate = "HTML Templates/Common Master-Detail.html";
                    const string strDocumentReport = "HTML Reports/Common Master-Detail Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // This is needed solely for images in HTML documents.
                        assembler.KnownTypes.Add(typeof(CommonUtilities.FileUtil));
                        // Call AssembleDocument to generate Common Master-Detail Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonMasterDetailinHtmlFormat
                }
                    break;
                case "text":
                {
                    //ExStart:GenerateCommonMasterDetailinTextFormat
                    const string strTxtDocumentTemplate = "Text Templates/Common Master-Detail.txt";
                    const string strTxtDocumentReport = "Text Reports/Common Master-Detail Report.txt";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Common Master-Detail Report in text document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strTxtDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strTxtDocumentReport),
                            new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonMasterDetailinTextFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateCommonMasterDetailinEmailFormat
                    const string strEmailDocumentTemplate = "Email Templates/Common Master-Detail.msg";
                    const string strEmailDocumentReport = "Email Reports/Common Master-Detail Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources =
                            DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.PopulateData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

                        // Call AssembleDocument to generate Common Master-Detail Report in email document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailDocumentReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateCommonMasterDetailinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateInParagraphList(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInParagraphListFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Paragraph List_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Paragraph List_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strDocumentTemplate = "Word Templates/In-Paragraph List_XML_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Paragraph List_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInParagraphListReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Paragraph List_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateInParagraphListinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Paragraph List Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Paragraph List_XML_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInParagraphListReportFromJsoninOpenExcel
                        const string strDocumentTemplate = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/In-Paragraph List_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListReportFromJsoninOpenExcel
                    }
                    else
                    {
                        //ExStart:GenerateInParagraphListinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List Report.ods";
                        
                        try
                        {  
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Paragraph List_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Paragraph List_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Paragraph List_XML_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Paragraph List_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInParagraphListReportFromJsoninOpenPresentation
                        const string strDocumentTemplate = "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/In-Paragraph List_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateInParagraphListinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Paragraph List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Paragraph List Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Paragraph List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInParagraphListinOpenPresentationFormat
                    }

                    break;
                case "html":
                {
                    //ExStart:GenerateInParagraphListinHtmlFormat
                    const string strDocumentTemplate = "HTML Templates/In-Paragraph List.html";
                    const string strDocumentReport = "HTML Reports/In-Paragraph List Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Paragraph List Report in htmlformat.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInParagraphListinHtmlFormat
                }
                    break;

                case "text":
                {
                    //ExStart:GenerateInParagraphListinTextFormat
                    const string strDocumentTemplate = "Text Templates/In-Paragraph List.txt";
                    const string strDocumentReport = "Text Reports/In-Paragraph List Report.txt";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Paragraph List Report in text document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInParagraphListinTextFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateInParagraphListinEmailFormat
                    const string strEmailDocumentTemplate = "Email Templates/In-Paragraph List.msg";
                    const string strEmailDocumentReport = "Email Reports/In-Paragraph List Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources =
                            DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.GetProductsData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

                        // Call AssembleDocument to generate In-Paragraph List Report in email document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailDocumentReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInParagraphListinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateInTableListWithAlternateContent(string strDocumentFormat, bool isDatabase,
            bool isDataSet, bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Alternate Content_DB_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Alternate Content_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Alternate Content_DT_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Alternate Content_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDt(), "ds"));
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
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Alternate Content_XML_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Alternate Content_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithAlternateContentReportFromJsoninOpenWord
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Alternate Content_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Alternate Content_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithAlternateContentReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithAlternateContentinOpenDocumentProcessingFormat
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Alternate Content_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List with Alternate Content Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Alternate Content_DB_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Alternate Content_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Alternate Content_DT_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Alternate Content_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDt(), "ds"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Alternate Content_XML_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Alternate Content_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithAlternateContentReportFromJsoninOpenExcel
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/In-Table List with Alternate Content_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithAlternateContentReportFromJsoninOpenExcel
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithAlternateContentinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Alternate Content Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Alternate Content_DB_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Alternate Content_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Alternate Content_DT_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Alternate Content_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDt(), "ds"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Alternate Content_XML_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Alternate Content_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithAlternateContentReportFromJsoninOpenPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/In-Table List with Alternate Content_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content report in presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithAlternateContentReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithAlternateContentinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Alternate Content Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithAlternateContentinOpenPresentationFormat
                    }

                    break;
                case "html":

                {
                    //ExStart:GenerateInTableListWithAlternateContentinHtmlFormat
                    const string strHtmlTemplate = "HTML Templates/In-Table List with Alternate Content.html";
                    const string strHtmlReport = "HTML Reports/In-Table List with Alternate Content Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Alternate Content Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                            CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithAlternateContentinHtmlFormat
                }
                    break;

                case "email":

                {
                    //ExStart:GenerateInTableListWithAlternateContentinEmailFormat
                    const string strEmailTemplate = "Email Templates/In-Table List with Alternate Content.msg";
                    const string strEmailReport = "Email Reports/In-Table List with Alternate Content Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

                        // Call AssembleDocument to generate In-Table List with Alternate Content Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithAlternateContentinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateInTableListWithFilteringGroupingAndOrdering(string strDocumentFormat,
            bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinDocumentProcessingDocument
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenWord
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Filtering, Grouping, and Ordering Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginOpenPresentationFormat
                    }

                    break;
                case "html":
                {
                    //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginHtmlDocument
                    const string strHtmlTemplate =
                        "HTML Templates/In-Table List with Filtering, Grouping, and Ordering.html";
                    const string strHtmlReport =
                        "HTML Reports/In-Table List with Filtering, Grouping, and Ordering Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                            CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginHtmlDocument
                }
                    break;

                case "email":
                {
                    //ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginEmailDocument
                    const string strEmailTemplate =
                        "Email Templates/In-Table List with Filtering, Grouping, and Ordering.msg";
                    const string strEmailReport =
                        "Email Reports/In-Table List with Filtering, Grouping, and Ordering Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

                        // Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginEmailDocument
                }
                    break;

            }
        }

        public static void GenerateInTableListWithHighlightedRows(string strDocumentFormat, bool isDatabase,
            bool isDataSet, bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Highlighted Rows_DB_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Highlighted Rows_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Highlighted Rows_DT_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Highlighted Rows_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Highlighted Rows_XML_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Highlighted Rows_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinDocumentProcessingDocument
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenWord
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                        const string strDocumentTemplate =
                            "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List with Highlighted Rows Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Highlighted Rows_DB_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Highlighted Rows_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // all AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Highlighted Rows_DT_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Highlighted Rows_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Highlighted Rows_XML_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Highlighted Rows_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinSpreadsheetDocument
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/In-Table List with Highlighted Rows Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Highlighted Rows_DB_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Highlighted Rows_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersDataDb(), "orders"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Highlighted Rows_DT_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Highlighted Rows_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Highlighted Rows_XML_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Highlighted Rows_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateInTableListWithHighlightedRowsinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table List with Highlighted Rows Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListWithHighlightedRowsinOpenPresentationFormat
                    }

                    break;

                case "html":
                {
                    //ExStart:GenerateInTableListWithHighlightedRowsinHtmlDocument
                    const string strHtmlTemplate = "HTML Templates/In-Table List with Highlighted Rows.html";
                    const string strHtmlReport = "HTML Reports/In-Table List with Highlighted Rows Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                            CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithHighlightedRowsinHtmlDocument
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateInTableListWithHighlightedRowsinEmailDocument
                    const string strEmailTemplate = "Email Templates/In-Table List with Highlighted Rows.msg";
                    const string strEmailReport = "Email Reports/In-Table List with Highlighted Rows Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

                        // Call AssembleDocument to generate In-Table List with Highlighted Rows Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithHighlightedRowsinEmailDocument
                }
                    break;

            }
        }

        public static void GenerateInTableList(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableListFromDatabaseinDocumentProcessingDocument
                        const string strDocumentTemplate = "Word Templates/In-Table List_DB_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strDocumentTemplate = "Word Templates/In-Table List_DT_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strDocumentTemplate = "Word Templates/In-Table List_XML_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromXMLinDocumentProcessingDocument
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/In-Table List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateInTableListinDocumentProcessingDocument
                        const string strDocumentTemplate = "Word Templates/In-Table List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table List Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_DB_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table List_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_DT_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table List_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table List_XML_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table List_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate = "Spreadsheet Templates/In-Table List_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/In-Table List_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateInTableListinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table List Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List_DB_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Table List_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List_DT_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Table List_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table List_XML_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Table List_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableListReportFromJsoninOpenPresentation
                        const string strDocumentTemplate = "Presentation Templates/In-Table List_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/In-Table List_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateInTableListinOpenPresentationFormat
                        const string strPresentationTemplate = "Presentation Templates/In-Table List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Table List Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableListinOpenPresentationFormat
                    }

                    break;

                case "html":
                {
                    //ExStart:GenerateInTableListinHtmlDocument
                    const string strHtmlTemplate = "HTML Templates/In-Table List.html";
                    const string strHtmlReport = "HTML Reports/In-Table List Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                            CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListinHtmlDocument
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateInTableListinEmailDocument
                    const string strEmailTemplate = "Email Templates/In-Table List.msg";
                    const string strEmailReport = "Email Reports/In-Table List Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.PopulateData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

                        // Call AssembleDocument to generate In-Table List Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListinEmailDocument
                }
                    break;
            }
        }

        public static void GenerateInTableMasterDetail(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateInTableMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/In-Table Master-Detail_DB_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table Master-Detail_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strDocumentTemplate = "Word Templates/In-Table Master-Detail_DT_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table Master-Detail_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strDocumentTemplate = "Word Templates/In-Table Master-Detail_XML_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table Master-Detail_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromXMLinOpenDocumentProcessingFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableMasterDetailReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/In-Table Master-Detail_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/In-Table Master-Detail_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateInTableMasterDetailinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/In-Table Master-Detail_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/In-Table Master-Detail Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table Master-Detail_DB_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table Master-Detail_DT_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table Master-Detail_XML_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableMasterDetailReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/In-Table Master-Detail_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateInTableMasterDetailinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table Master-Detail_DB_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table Master-Detail_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersDataDb(), "customers"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table Master-Detail_DT_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table Master-Detail_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDt(), "ds"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table Master-Detail_XML_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/In-Table Master-Detail_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateInTableMasterDetailReportFromJsoninOpenPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/In-Table Master-Detail_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/In-Table Master-Detail_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // all AssembleDocument to generate In-Table Master-Detail report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateInTableMasterDetailinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/In-Table Master-Detail_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/In-Table Master-Detail Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateInTableMasterDetailinOpenPresentationFormat
                    }

                    break;
                case "html":
                {
                    //ExStart:GenerateInTableMasterDetailinHtmlFormat
                    const string strHtmlTemplate = "HTML Templates/In-Table Master-Detail.html";
                    const string strHtmlReport = "HTML Reports/In-Table Master-Detail Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table Master-Detail Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                            CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.PopulateData(), "customers"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableMasterDetailinHtmlFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateInTableMasterDetailinEmailFormat
                    const string strEmailTemplate = "Email Templates/In-Table Master-Detail.msg";
                    const string strEmailReport = "Email Reports/In-Table Master-Detail Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.PopulateData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

                        // Call AssembleDocument to generate In-Table Master-Detail Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableMasterDetailinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateInTableListWithProgressiveTotal(string strDocumentFormat, bool isDatabase,
            bool isDataSet, bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:GenerateInTableListWithProgressiveTotalinOpenDocumentProcessingFormat
                    const string strDocumentTemplate =
                        "Word Templates/In-Table List with Running (Progressive) Total.docx";
                    const string strDocumentReport =
                        "Word Reports/In-Table List with Running (Progressive) Total Report.docx";
                    
                    try
                    {
                        var testData = DataLayer.GetOrdersData();
                        
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Progressive(Running) total Report in open document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithProgressiveTotalinOpenDocumentProcessingFormat

                    break;

                case "spreadsheet":
                {
                    //ExStart:GenerateInTableListWithProgressiveTotalinOpenSpreadsheetFormat
                    const string strSpreadsheetTemplate =
                        "Spreadsheet Templates/In-Table List with Running (Progressive) Total.xlsx";
                    const string strSpreadsheetReport =
                        "Spreadsheet Reports/In-Table List with Running (Progressive) Total Report.xlsx";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Progressive(Running) total Report in open spreadsheet format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                            CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithProgressiveTotalinOpenSpreadsheetFormat
                }
                    break;

                case "presentation":
                {
                    //ExStart:GenerateInTableListWithProgressiveTotalinOpenPresentationFormat
                    const string strPresentationTemplate =
                        "Presentation Templates/In-Table List with Running (Progressive) Total.pptx";
                    const string strPresentationReport =
                        "Presentation Reports/In-Table List with Running (Progressive) Total Report.pptx";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Progressive(running) total Report in open presentation format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                            CommonUtilities.SetDestinationDocument(strPresentationReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithProgressiveTotalinOpenPresentationFormat
                }
                    break;
                case "html":
                {
                    //ExStart:GenerateInTableListWithProgressiveTotalinHtmlFormat
                    const string strHtmlTemplate = "HTML Templates/In-Table List with Running (Progressive) Total.html";
                    const string strHtmlReport = "HTML Reports/In-Table List with Running (Progressive) Total.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate In-Table List with Progressive(Running) Total Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                            CommonUtilities.SetDestinationDocument(strHtmlReport),
                            new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithProgressiveTotalinHtmlFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateInTableListWithProgressiveTotalinEmailFormat
                    const string strEmailTemplate =
                        "Email Templates/In-Table List with Running (Progressive) Total.msg";
                    const string strEmailReport = "Email Reports/In-Table List with Running (Progressive) Total.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

                        // Call AssembleDocument to generate In-Table List with Progressive(Running) Total Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateInTableListWithProgressiveTotalinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateMulticoloredNumberedList(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenDocumentProcessingDocument
                        const string strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Multicolored Numbered List_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDatabaseinOpenDocumentProcessingDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDataSetinOpenDocumentProcessingDocument
                        const string strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Multicolored Numbered List_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDataSetinOpenDocumentProcessingDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromXMLinOpenDocumentProcessingDocument
                        const string strDocumentTemplate =
                            "Word Templates/Multicolored Numbered List_XML_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Multicolored Numbered List_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenDocumentProcessingDocument
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateMulticoloredNumberedListReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        const string strDocumentReport =
                            "Word Reports/Multicolored Numbered List_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateMulticoloredNumberedListinOpenDocumentProcessingDocument
                        const string strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Multicolored Numbered List Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListinOpenDocumentProcessingDocument
                    }

                    break;

                case "spreadsheet":
                    if (isDatabase)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenSpreadsheetDocument
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Multicolored Numbered List_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDatabaseinOpenSpreadsheetDocument
                    }
                    else if (isDataSet)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDataSetinOpenSpreadsheetDocument
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Multicolored Numbered List_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromDataSetinOpenSpreadsheetDocument
                    }
                    else if (isDataSourceXML)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromXMLinOpenSpreadsheetDocument
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Multicolored Numbered List_XML_OpenDocument.ods";
                        const string strSpreadsheetReport =
                            "Spreadsheet Reports/Multicolored Numbered List_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenSpreadsheetDocument
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateMulticoloredNumberedListReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate =
                            "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/Multicolored Numbered List_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateMulticoloredNumberedListinOpenSpreadsheetDocument
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListinOpenSpreadsheetDocument
                    }

                    break;

                case "presentation":
                    if (isDatabase)
                    {
                        //ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/Multicolored Numbered List_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/Multicolored Numbered List_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Multicolored Numbered List_XML_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/Multicolored Numbered List_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateMulticoloredNumberedListReportFromJsoninOpenPresentation
                        const string strDocumentTemplate =
                            "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/Multicolored Numbered List_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateMulticoloredNumberedListinOpenPresentationFormat
                        const string strPresentationTemplate =
                            "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
                        const string strPresentationReport =
                            "Presentation Reports/Multicolored Numbered List Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateMulticoloredNumberedListinOpenPresentationFormat
                    }

                    break;
                case "html":
                {
                    //ExStart:GenerateMulticoloredNumberedListinHtml
                    const string strDocumentTemplate = "HTML Templates/Multicolored Numbered List.html";
                    const string strDocumentReport = "HTML Reports/Multicolored Numbered List Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Multicolored Numbered List Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateMulticoloredNumberedListinHtml
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateMulticoloredNumberedListinEmail
                    const string strEmailDocumentTemplate = "Email Templates/Multicolored Numbered List.msg";
                    const string strEmailDocumentReport = "Email Reports/Multicolored Numbered List Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        
                        var dataSources =
                            DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.GetProductsData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

                        // Call AssembleDocument to generate Multicolored Numbered List Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailDocumentReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateMulticoloredNumberedListinEmail
                } break;
            }
        }

        public static void GenerateNumberedList(string strDocumentFormat, bool isDatabase, bool isDataSet,
            bool isDataSourceXML, bool isJson)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    if (isDatabase)
                    {
                        //ExStart:GenerateNumberedListFromDatabaseinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Numbered List_DB Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Numbered List_DT Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strDocumentTemplate = "Word Templates/Numbered List_XML_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Numbered List_XML Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromXMLinOpenDocumentProcessingFormat 
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateNumberedListReportFromJsoninOpenWord
                        const string strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Numbered List_OpenDocument_Json Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListReportFromJsoninOpenWord
                    }
                    else
                    {
                        //ExStart:GenerateNumberedListinOpenDocumentProcessingFormat
                        const string strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
                        const string strDocumentReport = "Word Reports/Numbered List Report.odt";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open document format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Numbered List_DB Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Numbered List_DT Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strSpreadsheetTemplate =
                            "Spreadsheet Templates/Numbered List_XML_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Numbered List_XML Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromXMLinOpenSpreadsheetFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateNumberedListReportFromJsoninOpenSpreadsheet
                        const string strDocumentTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        const string strDocumentReport =
                            "Spreadsheet Reports/Numbered List_OpenDocument_Json Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListReportFromJsoninOpenSpreadsheet
                    }
                    else
                    {
                        //ExStart:GenerateNumberedListinOpenSpreadsheetFormat
                        const string strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
                        const string strSpreadsheetReport = "Spreadsheet Reports/Numbered List Report.ods";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open spreadsheet format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                                CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
                        const string strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Numbered List_DB Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDataDb(), "products"));
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
                        const string strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Numbered List_DT Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsDt()));
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
                        const string strPresentationTemplate =
                            "Presentation Templates/Numbered List_XML_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Numbered List_XML Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListFromXMLinOpenPresentationFormat
                    }
                    else if (isJson)
                    {
                        //ExStart:GenerateNumberedListReportFromJsoninOpenPresentation
                        const string strDocumentTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        const string strDocumentReport =
                            "Presentation Reports/Numbered List_OpenDocument_Json Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                CommonUtilities.SetDestinationDocument(strDocumentReport),
                                new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListReportFromJsoninOpenPresentation
                    }
                    else
                    {
                        //ExStart:GenerateNumberedListinOpenPresentationFormat
                        const string strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
                        const string strPresentationReport = "Presentation Reports/Numbered List Report.odp";
                        
                        try
                        {
                            DocumentAssembler assembler = new DocumentAssembler();
                            // Call AssembleDocument to generate Numbered List Report in open presentation format.
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                                CommonUtilities.SetDestinationDocument(strPresentationReport),
                                new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        //ExEnd:GenerateNumberedListinOpenPresentationFormat
                    }

                    break;

                case "html":
                {
                    //ExStart:GenerateNumberedListinHtmlFormat
                    const string strDocumentTemplate = "HTML Templates/Numbered List.html";
                    const string strDocumentReport = "HTML Reports/Numbered List Report.html";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Numbered List Report in html format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateNumberedListinHtmlFormat
                }
                    break;
                case "text":
                {
                    //ExStart:GenerateNumberedListinTextFormat
                    const string strDocumentTemplate = "Text Templates/Numbered List.txt";
                    const string strDocumentReport = "Text Reports/Numbered List Report.txt";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate Numbered List Report in text document  format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetProductsData(), "products"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateNumberedListinTextFormat
                }
                    break;
                case "email":
                {
                    //ExStart:GenerateNumberedListinEmailFormat
                    const string strEmailDocumentTemplate = "Email Templates/Numbered List.msg";
                    const string strEmailDocumentReport = "Email Reports/Numbered List Report.msg";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();

                        var dataSources =
                            DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.GetProductsData());
                        var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

                        // Call AssembleDocument to generate Numbered List Report in email format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strEmailDocumentReport),
                            new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                            new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                            new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                            new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                            new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:GenerateNumberedListinEmailFormat
                }
                    break;
            }
        }

        public static void GenerateNumberedListRestartNum_Documents()
        {
            //ExStart:GenerateNumberedListRestartNum
            const string strDocumentTemplate = "Word Templates/Numbered List_RestartNum.docx";
            const string strDocumentReport = "Word Reports/Numbered List_RestartNum.docx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Numbered List report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:GenerateNumberedListRestartNum
        }


        public static void GenerateNumberedListRestartNum_Email()
        {
            //ExStart:GenerateNumberedListRestartNum
            const string strDocumentTemplate = "Email Templates/Numbered List_RestartNum.msg";
            const string strDocumentReport = "Email Reports/Numbered List_RestartNum.msg";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Numbered List report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:GenerateNumberedListRestartNum
        }

        public static void GenerateReportLazilyAndRecursively()
        {
            //ExStart:GeneratingReportbyRecursivelyandLazilyAccessingtheData
            const string strDocumentTemplate = "Word Templates/Lazy And Recursive.docx";
            const string strDocumentReport = "Word Reports/Lazy And Recursive Report.docx";
            
            try
            {
                DynamicEntity dentity = new DynamicEntity(Guid.NewGuid());
                
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Single Row Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(dentity, "root"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:GeneratingReportbyRecursivelyandLazilyAccessingtheData
        }

        public static void GenerateReportUsingMultipleDS(string strDocumentFormat)
        {
            if (strDocumentFormat == "document")
            {
                //ExStart:GeneratingReportUsingMultipleDataSourcesdocumentprocessing
                const string strDocumentTemplate = "Word Templates/Multiple DS.odt";
                const string strDocumentReport = "Word Reports/Multiple DS.odt";
                
                try
                {
                    DocumentAssembler assembler = new DocumentAssembler();
                    // Call AssembleDocument to generate report.
                    assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                        CommonUtilities.SetDestinationDocument(strDocumentReport),
                        new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"),
                        new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                //ExEnd:GeneratingReportUsingMultipleDataSourcesdocumentprocessing
            }
            else if (strDocumentFormat == "spreadsheet")
            {
                //ExStart:GeneratingReportUsingMultipleDataSourcesspreadsheet
                const string strDocumentTemplate = "Spreadsheet Templates/Multiple DS.ods";
                const string strDocumentReport = "Spreadsheet Reports/Multiple DS.ods";
                
                try
                {
                    DocumentAssembler assembler = new DocumentAssembler();
                    // Call AssembleDocument to generate report.
                    assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                        CommonUtilities.SetDestinationDocument(strDocumentReport),
                        new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"),
                        new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                //ExEnd:GeneratingReportUsingMultipleDataSourcesspreadsheet
            }
            else if (strDocumentFormat == "presentation")
            {
                //ExStart:GeneratingReportUsingMultipleDataSourcespresentation
                const string strDocumentTemplate = "Presentation Templates/Multiple DS.odp";
                const string strDocumentReport = "Presentation Reports/Multiple DS.odp";
                
                try
                {
                    object[] dataSourceObj = new object[]
                        {DataLayer.GetAllDataFromXml(), DataLayer.GetProductsDataJson()};
                    string[] dataSourceString = new string[] {"ds", "products"};

                    DocumentAssembler assembler = new DocumentAssembler();
                    // Call AssembleDocument to generate report.
                    assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                        CommonUtilities.SetDestinationDocument(strDocumentReport),
                        new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"),
                        new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                //ExEnd:GeneratingReportUsingMultipleDataSourcespresentation
            }

        }

        public static void TemplateSyntaxFormatting()
        {
            //ExStart:TemplateSyntaxFormatting 
            const string strDocumentTemplate = "Word Templates/String_Numeric_Formatting.odt";
            const string strDocumentReport = "Word Reports/String_Numeric_Formatting Report.odt";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate   Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TemplateSyntaxFormatting 
        }

        public static void OuterDocumentInsertion()
        {
            //ExStart:OuterDocumentInsertion
            const string strDocumentTemplate = "Word Templates/OuterDocInsertion.odt";
            const string strDocumentReport = "Word Reports/OuterDocInsertion Report.odt";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate  Report in open document format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:OuterDocumentInsertion
        }

        public static void AddBarCodes(string strDocumentFormat)
        {
            switch (strDocumentFormat)
            {
                case "document":
                    //ExStart:AddBarCodesDocumentProcessingFormat
                    const string strDocumentTemplate = "Word Templates/Barcode.docx";
                    const string strDocumentReport = "Word Reports/Barcode.docx";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate   Report in open document format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                            CommonUtilities.SetDestinationDocument(strDocumentReport),
                            new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:AddBarCodesDocumentProcessingFormat 
                    break;

                case "spreadsheet":
                    //ExStart:AddBarCodesSpreadsheet
                    const string strSpreadsheetTemplate = "Spreadsheet Templates/Barcode.xlsx";
                    const string strSpreadsheetReport = "Spreadsheet Reports/Barcode.xlsx";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate  Report in open spreadsheet format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate),
                            CommonUtilities.SetDestinationDocument(strSpreadsheetReport),
                            new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:AddBarCodesSpreadsheet 
                    break;

                case "presentation":
                    //ExStart:AddBarCodesPowerPoint
                    const string strPresentationTemplate = "Presentation Templates/Barcode.pptx";
                    const string strPresentationReport = "Presentation Reports/Barcode.pptx";
                    
                    try
                    {
                        DocumentAssembler assembler = new DocumentAssembler();
                        // Call AssembleDocument to generate  Report in open presentation format.
                        assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate),
                            CommonUtilities.SetDestinationDocument(strPresentationReport),
                            new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //ExEnd:AddBarCodesPowerPoint 
                    break;
            }
        }

        public static void UpdateWordDocFields(string documentType)
        {
            if (documentType == "document")
            {
                //ExStart:UpdateWordDocFields
                const string strDocumentTemplate = "Word Templates/Update_Field_XML.docx";
                const string strDocumentReport = "Word Reports/Update_Field_XML Report.docx";
                
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.Options |= DocumentAssemblyOptions.UpdateFieldsAndFormulas;
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                //ExEnd:UpdateWordDocFields
            }
            else if (documentType == "spreadsheet")
            {
                //ExStart:updateformula
                const string strDocumentTemplate = "Spreadsheet Templates/Update-Fomula.xlsx";
                const string strDocumentReport = "Spreadsheet Reports/Update-Fomula Report.xlsx";
                
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.Options |= DocumentAssemblyOptions.UpdateFieldsAndFormulas;
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
                //ExEnd:updateformula
            }
        }

        /// <summary>
        /// Support for Analogue of Microsoft Word NEXT Field.
        /// </summary>
        public static void NextIteration()
        {
            //ExStart:nextiteration
            const string strDocumentTemplate = "Word Templates/Using Next.docx";
            const string strDocumentReport = "Word Reports/Using Next Report.docx";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to generate Report.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:nextiteration
        }

        /// <summary>
        /// Generate report from excel data source.
        /// </summary>
        public static void UseSpreadsheetAsDataSource()
        {
            //ExStart:UseSpreadsheetAsDataSource
            string strDocumentTemplate = "Word Templates/Using Spreadsheet as Table of Data.docx";
            string strDocumentReport = "Word Reports/Using Spreadsheet as Table of Data_Output.docx";

            // Assemble a document using the external document table as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                CommonUtilities.SetDestinationDocument(strDocumentReport),
                new DataSourceInfo(DataLayer.ExcelData(), "contracts"));
            //ExEnd:UseSpreadsheetAsDataSource
        }

        /// <summary>
        /// Generate report from presentation data source.
        /// </summary>
        public static void UsePresentationTableAsDataSource()
        {
            string strDocumentTemplate = "Presentation Templates/Using Presentation as Table of Data.pptx";
            string strDocumentReport = "Presentation Reports/Using Presentation as Table of Data_Output.pptx";
            
            // Assemble a document using the external document table as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                CommonUtilities.SetDestinationDocument(strDocumentReport),
                new DataSourceInfo(DataLayer.PresentationData(), "table"));
        }

        /// <summary>
        /// Import word proccessing table into presentation.
        /// </summary>
        public static void ImportingWordProcessingTableIntoPresentation()
        {
            string strDocumentTemplate =
                "Presentation Templates/Importing Word Processing Table into Presentation.pptx";
            string strDocumentReport =
                "Presentation Reports/Importing Word Processing Table into Presentation_Output.pptx";


            // Assemble a document using the external document table as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                CommonUtilities.SetDestinationDocument(strDocumentReport),
                new DataSourceInfo(DataLayer.ImportingWordDocToPresentation(), "table"));
        }

        /// <summary>
        /// Import Spreadsheet into HTML Document.
        /// </summary>
        public static void ImportingSpreadsheetIntoHtmlDocument()
        {
            //ExStart:ImportingSpreadsheetIntoHtmlDocument
            try
            {
                string strHtmlTemplate = "HTML Templates/Importing Spreadsheet into HTML Document.html";
                string strHtmlReport = "HTML Reports/Importing Spreadsheet into HTML Document.html";


                // Assemble a document using the external document table as a data source.
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate),
                    CommonUtilities.SetDestinationDocument(strHtmlReport),
                    new DataSourceInfo(DataLayer.ImportingSpreadsheetToHtml(), "table"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:ImportingSpreadsheetIntoHtmlDocument
        }

        /// <summary>
        /// Table Cell Merging in Word Processing documents
        /// features is supported by version 19.1 or greater.
        /// </summary>
        public static void TableCellsMergingInWordProcessing()
        {
            //ExStart:TableCellsMergingInWordProcessing
            const string strDocumentTemplate = "Word Templates/Merging Cells Dynamically.docx";
            const string strDocumentReport = "PDF Reports/Merging Cells Dynamically Report.pdf";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to Merging Cells Dynamically Report in PDF format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInWordProcessing
        }

        /// <summary>
        /// Table Cell Merging in Presentations
        /// features is supported by version 19.1 or greater.
        /// </summary>
        public static void TableCellsMergingInPresentations()
        {
            //ExStart:TableCellsMergingInPresentations
            const string strDocumentTemplate = "Presentation Templates/Merging Cells Dynamically.pptx";
            const string strDocumentReport = "PDF Reports/Merging Cells Dynamically Report.pdf";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to Merging Cells Dynamically Report in PDF format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInPresentations
        }

        /// <summary>
        /// Table Cell Merging in Spreadsheets
        /// features is supported by version 19.1 or greater.
        /// </summary>
        public static void TableCellsMergingInSpreadsheets()
        {
            //ExStart:TableCellsMergingInSpreadsheets
            const string strDocumentTemplate = "Spreadsheet Templates/Merging Cells Dynamically.xlsx";
            const string strDocumentReport = "PDF Reports/Merging Cells Dynamically Report.pdf";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                // Call AssembleDocument to Merging Cells Dynamically Report in PDF format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInSpreadsheets
        }

        /// <summary>
        /// Table Cell Merging in Emails
        /// features is supported by version 19.1 or greater.
        /// </summary>
        public static void TableCellsMergingInEmails()
        {
            //ExStart:TableCellsMergingInEmails
            const string strEmailTemplate = "Email Templates/Merging Cells Dynamically.msg";
            const string strEmailReport = "Email Reports/Merging Cells Dynamically Report.msg";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.PopulateData());
                var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

                // Call AssembleDocument to generate Merging Cells Dynamically Report in email format.
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
                    CommonUtilities.SetDestinationDocument(strEmailReport),
                    new DataSourceInfo(dataSources.DataSource, dataSourcesNames.Name),
                    new DataSourceInfo(dataSources.Sender, dataSourcesNames.Sender),
                    new DataSourceInfo(dataSources.Recipients, dataSourcesNames.Recipients),
                    new DataSourceInfo(dataSources.CC, dataSourcesNames.CC),
                    new DataSourceInfo(dataSources.Subject, dataSourcesNames.Subject)
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInEmails
        }

        /// <summary>
        /// In-lining syntax error messages into templates
        /// features is supported by version 19.3 or greater.
        /// </summary>
        public static void DemoInLineSyntaxError()
        {
            //ExStart:DemoInLineSyntaxError_19.3
            const string strDocumentTemplate = "Word Templates/Inline Error Demo.docx";
            const string strDocumentReport = "PDF Reports/Inline Error Demo.pdf";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                // Enable the In-line error messaging.
                assembler.Options |= DocumentAssemblyOptions.InlineErrorMessages;

                // Call AssembleDocument to show the demo Report and save the report in PDF format.
                // The AssembleDocument will return a boolean value to indicate the success or failed with inline error.
                if (assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer")))
                {
                    Console.WriteLine("No error found in template");
                }
                else
                {
                    Console.WriteLine("Do something with a report containing a template syntax error.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:DemoInLineSyntaxError_19.3
        }

        /// <summary>
        /// Loading of template documents from HTML with resources.
        /// Features is supported by version 19.5 or greater.
        /// </summary>
        public static void LoadDocFromHTMLWithResource()
        {
            //ExStart:LoadDocFromHTMLWithResource_19.5
            string strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceLoad/");

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                // Assemble the document using relative directory path and file names.
                assembler.AssembleDocument(strDirectoryPath + "TestWordsResourceLoad.htm",
                    strDirectoryPath + "TestWordsResourceLoad_Out.docx",
                    new DataSourceInfo("It should be a jeep image.", "value"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:LoadDocFromHTMLWithResource_19.5
        }

        /// <summary>
        /// Loading of template documents from HTML with resources from an explicitly specified folder.
        /// Features is supported by version 19.5 or greater.
        /// </summary>
        public static void LoadDocFromHTMLWithResource_ExplicitFolder()
        {
            //ExStart:LoadDocFromHTMLWithResource_ExplicitFolder_19.5
            string strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceLoad/");

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                LoadSaveOptions loadSaveOptions = new LoadSaveOptions();

                // Resolve URIs from the specified alternative folder.
                loadSaveOptions.ResourceLoadBaseUri = strDirectoryPath + "Alternative";

                assembler.AssembleDocument(strDirectoryPath + "TestWordsResourceLoad.htm",
                    strDirectoryPath + "TestWordsResourceLoad_Out.docx", loadSaveOptions,
                    new DataSourceInfo("It should be a sport car image.", "value"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:LoadDocFromHTMLWithResource_ExplicitFolder_19.5
        }

        /// <summary>
        /// Saving of external resource files at relative path while an assembled document loaded from a non-HTML format is being saved to HTML.
        /// Features is supported by version 19.5 or greater.
        /// </summary>
        public static void SaveDocToHTMLWithResource()
        {
            //ExStart:SaveDocToHTMLWithResource_19.5
            string strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceSave/");
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                // Assemble the document using relative directory path and file names.
                assembler.AssembleDocument(strDirectoryPath + "TestCellsResourceSaveOneSheet.xlsx",
                    strDirectoryPath + "TestCellsResourceSaveOneSheet_Out.htm", new DataSourceInfo("Hello!", "value"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveDocToHTMLWithResource_19.5
        }

        /// <summary>
        /// Saving of external resource files in a specified folder at relative path while saving output to HTML.
        /// Features is supported by version 19.5 or greater.
        /// </summary>
        public static void SaveDocToHTMLWithResource_ExplicitFolder()
        {
            //ExStart:SaveDocToHTMLWithResource_ExplicitFolder_19.5
            string strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceSave/");

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                LoadSaveOptions loadSaveOptions = new LoadSaveOptions
                {
                    // Resolve URIs from the specified alternative folder.
                    ResourceSaveFolder = strDirectoryPath + "Alternative"
                };

                assembler.AssembleDocument(strDirectoryPath + "TestWordsResourceSave.docx",
                    strDirectoryPath + "TestWordsResourceSave_Out.htm", loadSaveOptions,
                    new DataSourceInfo("Hello!", "value"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveDocToHTMLWithResource_ExplicitFolder_19.5
        }

        /// <summary>
        /// Saving an assembled Markdown document to a Word Processing format using file extension.
        /// Features is supported by version 19.8 or greater.
        /// </summary>
        public static void SaveMdtoWord_UsingExtension(string DocumentName)
        {
            //ExStart:SaveMdtoWord_UsingExtension_19.8
            string strDocumentTemplate = "Markdown Templates/" + DocumentName;
            string strDocumentReport = "Word Reports/" + DocumentName + " Out.docx";
            const string description =
                "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                "office and email file formats based upon template documents and data obtained from various sources " +
                "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"),
                    new DataSourceInfo(description, "description"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveMdtoWord_UsingExtension_19.8
        }

        /// <summary>
        /// Saving an assembled Word Processing document or email to Markdown using file extension.
        /// Features is supported by version 19.8 or greater.
        /// </summary>
        public static void SaveWordOrEmailtoMD_UsingExtension()
        {
            //ExStart:SaveWordtoMD_UsingExtension_19.8
            const string strDocumentTemplate = "Word Templates/ReadMe.docx";
            const string strDocumentReport = "Markdown Reports/ReadMe Out.md";
            const string description =
                "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                "office and email file formats based upon template documents and data obtained from various sources " +
                "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"),
                    new DataSourceInfo(description, "description"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveWordtoMD_UsingExtension_19.8
        }

        /// <summary>
        /// Saving an assembled Word Processing document or email to Markdown using explicit specifying.
        /// Features is supported by version 19.8 or greater.
        /// </summary>
        public static void SaveWordOrEmailtoMD_Explicit()
        {
            //ExStart:SaveWordtoMD_Explicit_19.8
            try
            {
                Stream sourceStream = File.OpenRead(CommonUtilities.GetSourceDocument("Word Templates/ReadMe.docx"));
                FileStream targetStream =
                    File.Create(CommonUtilities.SetDestinationDocument("Markdown Reports/ReadMe Out.md"));

                const string description =
                    "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                    "office and email file formats based upon template documents and data obtained from various sources " +
                    "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

                DataSourceInfo dataSourceInfo1 =
                    new DataSourceInfo("GroupDocs.Assembly for .NET(Using Stream)", "product");
                DataSourceInfo dataSourceInfo2 = new DataSourceInfo(description, "description");

                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(sourceStream, targetStream, new LoadSaveOptions(FileFormat.Markdown),
                    dataSourceInfo1, dataSourceInfo2);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveWordtoMD_UsingExtension_19.8
        }

        /// <summary>
        /// Working with JSON data sources.
        /// Features is supported by version 19.10 or greater.
        /// </summary>
        public static void SimpleJsonDS_Demo()
        {
            //ExStart:SimpleJsonDS_Demo_19.10
            try
            {
                const string strDocumentTemplate = "Word Templates/SimpleDatasetDemo.docx";
                const string strDocumentReport = "Word Reports/SimpleJsonDSDemo Out.docx";
                const string strDataSource = "JSON DataSource/ManagerData.json";

                JsonDataSource dataSource = new JsonDataSource(CommonUtilities.GetDataSourceDocument(strDataSource));

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(dataSource, "managers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SimpleJsonDS_Demo_19.10
        }

        /// <summary>
        /// Working with XML data sources.
        /// Features is supported by version 19.10 or greater.
        /// </summary>
        public static void SimpleXMLDS_Demo()
        {
            //ExStart:SimpleXMLDS_Demo_19.10
            try
            {
                const string strDocumentTemplate = "Word Templates/SimpleDatasetDemo.docx";
                const string strDocumentReport = "Word Reports/SimpleXMLDSDemo Out.docx";
                const string strDataSource = "XML DataSource/Managers.xml";
                
                XmlDataSource dataSource = new XmlDataSource(CommonUtilities.GetDataSourceDocument(strDataSource));

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(dataSource, "managers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SimpleXMLDS_Demo_19.10
        }

        /// <summary>
        /// Working with csv data sources.
        /// Features is supported by version 19.10 or greater.
        /// </summary>
        public static void SimpleCsvDS_Demo()
        {
            //ExStart:SimpleCsvDS_Demo_19.10
            try
            {
                const string strDocumentTemplate = "Text Templates/CsvDatasetDemo.txt";
                const string strDocumentReport = "Word Reports/SimpleCsvDSDemo Out.docx";
                const string strDataSource = "Excel DataSource/Person.csv";

                CsvDataLoadOptions options = new CsvDataLoadOptions(true);
                CsvDataSource dataSource =
                    new CsvDataSource(CommonUtilities.GetDataSourceDocument(strDataSource), options);

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(dataSource, "persons"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SimpleCsvDS_Demo_19.10
        }

        /// <summary>
        /// Insert Image Dynamically in Word Document.
        /// Features is supported by version 20.3 or greater.
        /// </summary>
        public static void InsertImageDynamicallyInWord()
        {
            //ExStart:InsertImageDynamicallyInWord_20.3
            try
            {
                const string strDocumentTemplate = "Word Templates/DynamicImageDemo.docx";
                const string strDocumentReport = "Word Reports/DynamicImageDemo Out.docx";

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(CommonUtilities.GetImageFolder("no-photo.jpg"), "image_expression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:InsertImageDynamicallyInWord_20.3
        }

        /// <summary>
        /// Insert Document Dynamically in Word Document.
        /// Features is supported by version 20.3 or greater.
        /// </summary>
        public static void InsertDocumentDynamicallyInWord()
        {
            //ExStart:InsertDocumentDynamicallyInWord_20.3
            try
            {
                const string strDocumentTemplate = "Word Templates/DynamicDocInsert.docx";
                const string strDocumentReport = "Word Reports/DynamicDocInsert Out.docx";

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(CommonUtilities.GetOuterDocumentFolder("OuterDocument.docx"),
                        "document_expression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:InsertDocumentDynamicallyInWord_20.3
        }


        /// <summary>
        /// Set checkbox value dynamically in Word document.
        /// Features is supported by version 20.3 or greater.
        /// </summary>
        public static void SetCheckboxValueDynamicallyInWord(bool boolVal)
        {
            //ExStart:SetCheckboxValueDynamicallyInWord_20.3
            try
            {
                const string strDocumentTemplate = "Word Templates/CheckBoxValueSetDemo.docx";
                const string strDocumentReport = "Word Reports/CheckBoxValueSetDemo Out.docx";

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo(boolVal, "conditional_expression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SetCheckboxValueDynamicallyInWord_20.3
        }

        /// <summary>
        /// Saving template POT presentation document to assembled Presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        public static void SavePOTtoPPTX()
        {
            //ExStart:SavePOTtoPPTX_20.6
            const string strDocumentTemplate = "Presentation Templates/template.pot";
            const string strDocumentReport = "Presentation Reports/template.pot Out.pptx";

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SavePOTtoPPTX_20.6
        }

        /// <summary>
        /// Saving template OTP presentation document to assembled Presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        public static void SaveOTPtoPPTX()
        {
            //ExStart:SaveOTPtoPPTX_20.6
            const string strDocumentTemplate = "Presentation Templates/template.otp";
            const string strDocumentReport = "Presentation Reports/template.otp Out.pptx";

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveOTPtoPPTX_20.6
        }

        /// <summary>
        /// Saving assembled Presentation document to template OTP presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        public static void SavePPTXtoOTP()
        {
            //ExStart:SavePPTXtoOTP_20.6
            const string strDocumentTemplate = "Presentation Templates/template.pptx";
            const string strDocumentReport = "Presentation Reports/template.pptx Out.otp";

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SavePPTXtoOTP_20.6
        }

        /// <summary>
        /// Saving assembled Presentation document to template OTP presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        public static void SavePPTXtoPOT()
        {
            //ExStart:SavePPTXtoPOT_20.6
            const string strDocumentTemplate = "Presentation Templates/template.pptx";
            const string strDocumentReport = "Presentation Reports/template.pptx Out.pot";

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(
                    CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SavePPTXtoPOT_20.6
        }

        /// <summary>
        /// Saving assembled Presentation document to template OTP presentation document (stream).
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        public static void SavePPTXtoPOTOTPAsStream()
        {
            //ExStart:SavePPTXtoPOTAsStream_20.6
            Stream templateStream = new FileStream(CommonUtilities.GetSourceDocument("Presentation Templates/template.pptx"), FileMode.Open);
            Stream resultPotStream = new FileStream(CommonUtilities.SetDestinationDocument("Presentation Reports/template.pptx Out.pot"), FileMode.CreateNew);
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(templateStream, resultPotStream, new LoadSaveOptions(FileFormat.Pot),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SavePPTXtoPOTAsStream_20.6

            //ExStart:SavePPTXtoOTPAsStream_20.6
            templateStream.Seek(0, SeekOrigin.Begin);
            Stream resultOtpStream = new FileStream(CommonUtilities.SetDestinationDocument("Presentation Reports/template.pptx Out.otp"), FileMode.CreateNew);

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(templateStream, resultOtpStream, new LoadSaveOptions(FileFormat.Otp),
                    new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SavePPTXtoOTPAsStream_20.6
        }

        /// <summary>
        /// Changing the resulution of barcode images while saving the document.
        /// Uniform resolution in DPI across both the X and Y axes is supported by version 20.8 or greater.
        /// </summary>
        //ExStart:SetBarcodeResolution_20.8
        public static void SetBarcodeResolution()
        {
            const string strDocumentTemplate = "Word Templates/Barcode.docx";
            string strDocumentReport = "Word Reports/Barcode.72dpi.docx";

            AssembleDocumentSetBarcodeResolution(72, strDocumentTemplate, strDocumentReport);

            strDocumentReport = "Word Reports/Barcode.1600dpi.docx";

            AssembleDocumentSetBarcodeResolution(1600, strDocumentTemplate, strDocumentReport);
        }

        private static void AssembleDocumentSetBarcodeResolution(float resolution, string sourceTemplateFilename, string outputReportFilename)
        {
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.BarcodeSettings.Resolution = resolution;

                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(sourceTemplateFilename),
                    CommonUtilities.SetDestinationDocument(outputReportFilename));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        //ExEnd:SetBarcodeResolution_20.8

        /// <summary>
        /// Changing the scaling of the barcode image within its containing shape while saving the document.
        /// </summary>
        public static void SetBarcodeScale()
        {
            //ExStart:SetBarcodeScale_20.8
            const string strDocumentTemplate = "Word Templates/Barcode.docx";
            string strDocumentReport = "Word Reports/Barcode.docx";

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                assembler.BarcodeSettings.BaseXDimension *= 0.5f;
                assembler.BarcodeSettings.BaseYDimension *= 0.5f;

                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SetBarcodeScale_20.8
        }
    }
}