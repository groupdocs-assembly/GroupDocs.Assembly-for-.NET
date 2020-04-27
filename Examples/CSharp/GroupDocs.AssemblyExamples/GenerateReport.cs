using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.Assembly;
using GroupDocs.AssemblyExamples.BusinessLayer;
using System.Diagnostics;
using GroupDocs.AssemblyExamples.ProjectEntities;
using GroupDocs.Assembly.Data;
using System.IO;

namespace GroupDocs.AssemblyExamples
{
	public static class GenerateReport
	{
		/// <summary>
		/// Shows how to load document Table set using default options
		/// Features is supported by version 17.01 or greater
		/// </summary>
		/// <param name="dataSource">name of the data source file</param>
		public static void LoadDocTableSet(string dataSource)
		{
			//ExStart:LoadDocTableSet
			// Load all document tables using default options.
			DocumentTableSet tableSet = new DocumentTableSet(CommonUtilities.GetDataSourceDocument("Word DataSource/" + dataSource));

			// Check loading.
			Debug.Assert(tableSet.Tables.Count == 3);
			Debug.Assert(tableSet.Tables[0].Name == "Table1");
			Debug.Assert(tableSet.Tables[1].Name == "Table2");
			Debug.Assert(tableSet.Tables[2].Name == "Table3");
			//ExEnd:LoadDocTableSet
		}
		/// <summary>
		///Chang2 target file format using explicit specifying
		/// Features is supported by version 18.9 or greater
		/// </summary>
		public static void ChangeTargetFileFormatUsingExplicitSpecifying()
		{
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/Bubble Chart.docx";
			//Setting up destination PDF report 
			const String strDocumentReport = "PDF Reports/Bubble Chart Report.pdf";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to generate Bubble Chart Report in open document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport),new LoadSaveOptions(FileFormat.Pdf), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Change target file format using the file extension
		/// Features is supported by version 18.9 or greater
		/// </summary>
		public static void ChangeTargetFileFormat()
		{
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/Bubble Chart.docx";
			//Setting up destination PDF report 
			const String strDocumentReport = "PDF Reports/Bubble Chart Report.pdf";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to generate Bubble Chart Report in open document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
			DocumentTableSet tableSet = new DocumentTableSet(CommonUtilities.GetDataSourceDocument("Word DataSource/" + dataSource), new CustomDocumentTableLoadHandler());

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
		/// Shows how to use document TableSet as DataSource
		/// Features is supported by version 17.01 or greater
		/// </summary>
		/// <param name="dataSource">Name of the data source file</param>
		/// <param name="slideDoc">name of the template file</param>
		public static void UseDocumentTableSetAsDataSource(string dataSource, string slideDoc)
		{
			//ExStart:UseDocumentTableSetAsDataSource
			//setting up output document
			const string outDocument = "Presentation Reports/Use Document Table Set As DataSource Output.pptx";
			//set up path for the template file
			string templateFile = CommonUtilities.GetSourceDocument("Presentation Templates/" + slideDoc);
			// Set table column names to be extracted from the document.
			DocumentTableSet tableSet = new DocumentTableSet(CommonUtilities.GetDataSourceDocument("Word DataSource/" + dataSource), new ColumnNameExtractingDocumentTableLoadHandler());

			// Set table names for conveniency.
			tableSet.Tables[0].Name = "Planets";
			tableSet.Tables[1].Name = "Persons";
			tableSet.Tables[2].Name = "Companies";

			// Pass DocumentTableSet as a data source.
			DocumentAssembler assembler = new DocumentAssembler();
			assembler.AssembleDocument(templateFile, CommonUtilities.SetDestinationDocument(outDocument), new DataSourceInfo(tableSet));
			//ExEnd:UseDocumentTableSetAsDataSource
		}

		/// <summary>
		/// Shows how to define document table relations
		/// Feature is supported by version 17.01 or greater
		/// </summary>
		/// <param name="relatedTables">name of the data source file</param>
		/// <param name="docTableRelations">name of the template file</param>
		public static void DefiningDocumentTableRelations(string relatedTables, string docTableRelations)
		{
			//ExStart:DefiningDocumentTableRelations
			//setting up output document
			const string outDocument = "Word Reports/document relations output.docx";
			//set up path for the related tables data source
			string relatedTablesDataSource = CommonUtilities.GetDataSourceDocument("Excel DataSource/" + relatedTables);
			//set up path for the template file
			string templateFile = CommonUtilities.GetSourceDocument("Word Templates/" + docTableRelations);

			// Set table column names to be extracted from the document.
			DocumentTableSet tableSet = new DocumentTableSet(relatedTablesDataSource, new ColumnNameExtractingDocumentTableLoadHandler());

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
			assembler.AssembleDocument(templateFile, CommonUtilities.SetDestinationDocument(outDocument), new DataSourceInfo(tableSet));
			//ExEnd:DefiningDocumentTableRelations
		}

		/// <summary>
		/// Shows how to change document table column type
		/// Feature is supported by version 17.01 or greater
		/// </summary>
		/// <param name="document"></param>
		public static void ChangingDocumentTableColumnType(string document)
		{
			//ExStart:ChangingDocumentTableColumnType
			//setting up data source document
			const string dataSrcDocument = "Word DataSource/Managers Data.docx";
			//setting up output document
			const string outDocument = "Presentation Reports/Out.pptx";

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
			assembler.AssembleDocument(CommonUtilities.GetSourceDocument(document), CommonUtilities.SetDestinationDocument(outDocument), new DataSourceInfo(table, "Managers"));
			//ExEnd:ChangingDocumentTableColumnType
		}

		public static void GenerateBubbleChart(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Bubble Chart.docx";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Bubble Chart_Json Report.docx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Bubble Chart Report in document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Bubble Chart.docx";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Bubble Chart Report.docx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bubble Chart Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Bubble Chart.xlsx";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Bubble Chart_Json Report.xlsx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Bubble Chart Report in spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Bubble Chart.xlsx";
						//Setting up destination open document report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Bubble Chart Report.xlsx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Bubble Chart.pptx";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Bubble Chart_Json Report.pptx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Bubble Chart Report in presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strPresentationTemplate = "Presentation Templates/Bubble Chart.pptx";
						//Setting up destination open document report 
						const String strPresentationReport = "Presentation Reports/Bubble Chart Report.pptx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bubble Chart Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/Bubble Chart.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/Bubble Chart Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bubble Chart Report in open presentation format

							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

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
						//ExEnd:GenerateBubbleChartinEmailFormat
					}
					break;
			}
		}
        /// <summary>
        /// This method inertes nested external doucments in email document
        /// </summary>
        public static void InsertNestedExternalDocumentsInEmail()
        {
            //Setting up source open document template
            const String strDocumentTemplate = "Email Templates/Nested External Document.msg";
            //Setting up destination open document report 
            const String strDocumentReport = "Email Reports/Nested External Document.msg";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate  Report in open document format
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
        /// This method inertes nested external doucments in word document
        /// </summary>
        public static void InsertNestedExternalDocumentsInWord()
        {
            //Setting up source open document template
            const String strDocumentTemplate = "Word Templates/Nested External Document.docx";
            //Setting up destination open document report 
            const String strDocumentReport = "Word Reports/Nested External Document.docx";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate  Report in open document format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                           CommonUtilities.SetDestinationDocument(strDocumentReport),
                                           new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static void EmptyParagraphInEmail()
        {
            //setting up source 
            const String strDocumentTemplate = "Email Templates/Empty Paragraph.msg";
            //Setting up destination 
            const String strDocumentReport = "Email Reports/Empty Paragraph.msg";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler
                {
                    //Apply Remove Empty Paragraph option
                    Options = DocumentAssemblyOptions.RemoveEmptyParagraphs
                };
                //Call AssembleDocument to assemble document 
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                                           CommonUtilities.SetDestinationDocument(strDocumentReport),
                                           new DataSourceInfo("dummy", "dummy"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void EmptyParagraphInPresentation()
        {
            //setting up source 
            const String strDocumentTemplate = "Presentation Templates/Empty Paragraph.pptx";
            //Setting up destination 
            const String strDocumentReport = "Presentation Reports/Empty Paragraph.pptx";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler
                {
                    //Apply Remove Empty Paragraph option
                    Options = DocumentAssemblyOptions.RemoveEmptyParagraphs
                };
                //Call AssembleDocument to assemble document 
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
        /// Supports removal of paragraphs becoming empty after template syntax tags are removed or replaced with empty values for Word Processing documents s
        /// </summary>
        public static void EmptyParagraphInWordProcessing()
        {
            //setting up source 
            const String strDocumentTemplate = "Word Templates/Empty Paragraph.docx";
            //Setting up destination 
            const String strDocumentReport = "Word Reports/Empty Paragraph.docx";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler
                {
                    //Apply Remove Empty Paragraph option
                    Options = DocumentAssemblyOptions.RemoveEmptyParagraphs
                };
                //Call AssembleDocument to assemble document 
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
        /// Insert Hyperlink Dynamically in Email Document
        /// Feature is supported by version 18.7 or greater
        /// </summary>
        public static void DynamicHyperlinkInsertionEmail()
		{
			//setting up source 
			const String strDocumentTemplate = "Email Templates/Dynamic Hyperlink.msg";
			//Setting up destination 
			const String strDocumentReport = "Email Reports/Dynamic Hyperlink.msg";
			//Setting up Uri Expression
			const String uriExpression = "https://www.groupdocs.com/";
			//Setting up Display Text Expression
			const String displayTextExpression = "GroupDocs";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to assemble document 
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(uriExpression, "uriExpression"), new DataSourceInfo(displayTextExpression, "displayTextExpression"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
		/// <summary>
		/// Insert Hyperlink Dynamically in Spreadsheet Document
		/// Feature is supported by version 18.7 or greater
		/// </summary>
		public static void DynamicHyperlinkInsertionSpreadsheet()
		{
			//setting up source 
			const String strDocumentTemplate = "Spreadsheet Templates/Dynamic Hyperlink.xlsx";
			//Setting up destination 
			const String strDocumentReport = "Spreadsheet Reports/Dynamic Hyperlink.xlsx";
			//Setting up Uri Expression
			const String uriExpression = "https://www.groupdocs.com/";
			//Setting up Display Text Expression
			const String displayTextExpression = "GroupDocs";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to assemble document 
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(uriExpression, "uriExpression"), new DataSourceInfo(displayTextExpression, "displayTextExpression"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
		/// <summary>
		/// Insert Hyperlink Dynamically in Presentaion Document
		/// Feature is supported by version 18.7 or greater
		/// </summary>
		public static void DynamicHyperlinkInsertionPresentation()
		{
			//setting up source 
			const String strDocumentTemplate = "Presentation Templates/Dynamic Hyperlink.pptx";
			//Setting up destination 
			const String strDocumentReport = "Presentation Reports/Dynamic Hyperlink.pptx";
			//Setting up Uri Expression
			const String uriExpression = "https://www.groupdocs.com/";
			//Setting up Display Text Expression
			const String displayTextExpression = "GroupDocs";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to assemble document 
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(uriExpression, "uriExpression"), new DataSourceInfo(displayTextExpression, "displayTextExpression"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Insert Hyperlink Dynamically in Word Document
		/// Feature is supported by version 18.7 or greater
		/// </summary>
		public static void DynamicHyperlinkInsertionWord()
		{
			//setting up source 
			const String strDocumentTemplate = "Word Templates/Dynamic Hyperlink.docx";
			//Setting up destination 
			const String strDocumentReport = "Word Reports/Dynamic Hyperlink.docx";
			//Setting up Uri Expression
			const String uriExpression = "https://www.groupdocs.com/";
			//Setting up Display Text Expression
			const String displayTextExpression = "GroupDocs";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to assemble document 
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(uriExpression, "uriExpression"), new DataSourceInfo(displayTextExpression, "displayTextExpression"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
        /// <summary>
		/// Insert Bookmarks Dynamically in Word Document
		/// Feature is supported by version 20.1 or greater
		/// </summary>
		public static void DynamicBookmarkInsertionWord()
        {
            //setting up source 
            const String strDocumentTemplate = "Word Templates/Dynamic bookmarks.docx";
            //Setting up destination 
            const String strDocumentReport = "Word Reports/Dynamic bookmarks.docx";
            //Setting up Uri Expression
            const String bookmark_expression = "gd_bookmark";
            //Setting up Display Text Expression
            const String displayTextExpression = "GroupDocs";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to assemble document 
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), 
                    new DataSourceInfo(bookmark_expression, "bookmark_expression"), new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        /// <summary>
		/// Insert Bookmarks Dynamically in Excel Document
		/// Feature is supported by version 20.1 or greater
		/// </summary>
		public static void DynamicBookmarkInsertionExcel()
        {
            //setting up source 
            const String strDocumentTemplate = "Spreadsheet Templates/Dynamic Cell Range.xlsx";
            //Setting up destination 
            const String strDocumentReport = "Spreadsheet Reports/Dynamic Cell Range.xlsx";
            //Setting up Uri Expression
            const String bookmark_expression = "gd_bookmark";
            //Setting up Display Text Expression
            const String displayTextExpression = "GroupDocs";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to assemble document 
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(bookmark_expression, "bookmark_expression"), new DataSourceInfo(displayTextExpression, "displayTextExpression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        /// <summary>
        /// Sets colors of chart series point color dynamically based upon expressions
        /// Feature is supported by version 18.6 or greater
        /// </summary>
        public static void DynamicChartSeriesPointColorEmail()
		{
			//setting up source 
			const String strDocumentTemplate = "Email Templates/Dynamic Chart Point Series Color.msg";
			//Setting up destination 
			const String strDocumentReport = "Email Reports/Dynamic Chart Point Series Color.msg";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																	  //Call AssembleDocument to generate Pie Chart report in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Sets colors of chart series color dynamically based upon expressions
		/// Feature is supported by version 18.6 or greater
		/// </summary>
		public static void DynamicChartSeriesColorEmail()
		{
			// Setting up source open document template
			const String strDocumentTemplate = "Email Templates/Dynamic Chart Series Color.msg";
			//setting up data source document
			const string dataSrcDocument = "Word DataSource/Managers Data.docx";
			//Setting up destination open document report 
			const String strDocumentReport = "Email Reports/Dynamic Chart Series Color.msg";
			//Define serires color
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
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(table, "managers"), new DataSourceInfo(color, "color"));

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}

		}

		/// <summary>
		/// Sets colors of chart series point color dynamically based upon expressions
		/// Feature is supported by version 18.5 or greater
		/// </summary>
		public static void DynamicChartSeriesPointColorPresentation()
		{
			//setting up source 
			const String strDocumentTemplate = "Presentation Templates/Dynamic Chart Point Series Color.pptx";
			//Setting up destination 
			const String strDocumentReport = "Presentation Reports/Dynamic Chart Point Series Color.pptx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																	  //Call AssembleDocument to generate Pie Chart report in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Sets colors of chart series point color dynamically based upon expressions
		/// Feature is supported by version 18.5 or greater
		/// </summary>
		public static void DynamicChartSeriesPointColorSpreadsheet()
		{
			//setting up source 
			const String strDocumentTemplate = "Spreadsheet Templates/Dynamic Chart Point Series Color.xlsx";
			//Setting up destination 
			const String strDocumentReport = "Spreadsheet Reports/Dynamic Chart Point Series Color.xlsx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																	  //Call AssembleDocument to generate Pie Chart report in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Sets colors of chart series point color dynamically based upon expressions
		/// Feature is supported by version 18.5 or greater
		/// </summary>
		public static void DynamicChartSeriesPointColor()
		{
			//setting up source 
			const String strDocumentTemplate = "Word Templates/Dynamic Chart Point Series Color.docx";
			//Setting up destination 
			const String strDocumentReport = "Word Reports/Dynamic Chart Point Series Color.docx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																	  //Call AssembleDocument to generate Pie Chart report in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Sets colors of chart series dynamically based upon expressions
		/// Feature is supported by version 18.5 or greater
		/// </summary>
		public static void DynamicChartSeriesColorPresentation()
		{
			try
			{
				// Setting up source open document template
				const String strDocumentTemplate = "Presentation Templates/Dynamic Chart Series Color.pptx";
				//setting up data source document
				const string dataSrcDocument = "Presentation DataSource/Managers Data.pptx";
				//Setting up destination open document report 
				const String strDocumentReport = "Presentation Reports/Dynamic Chart Series Color.pptx";
				//Define serires color
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
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(table, "managers"), new DataSourceInfo(color, "color"));

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Sets colors of chart series dynamically based upon expressions
		/// Feature is supported by version 18.5 or greater
		/// </summary>
		public static void DynamicChartSeriesColorSpreadsheet()
		{
			try
			{
				// Setting up source open document template
				const String strDocumentTemplate = "Spreadsheet Templates/Dynamic Chart Series Color.xlsx";
				//setting up data source document
				const string dataSrcDocument = "Word DataSource/Managers Data.docx";
				//Setting up destination open document report 
				const String strDocumentReport = "Spreadsheet Reports/Dynamic Chart Series Color.xlsx";
				//Define serires color
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
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(table, "managers"), new DataSourceInfo(color, "color"));

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Sets colors of chart series dynamically based upon expressions
		/// Feature is supported by version 18.5 or greater
		/// </summary>
		public static void DynamicChartSeriesColor()
		{
			try
			{
				// Setting up source open document template
				const String strDocumentTemplate = "Word Templates/Dynamic Chart Series Color.docx";
				//setting up data source document
				const string dataSrcDocument = "Word DataSource/Managers Data.docx";
				//Setting up destination open document report 
				const String strDocumentReport = "Word Reports/Dynamic Chart Series Color.docx";
				//Define serires color
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
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(table, "managers"), new DataSourceInfo(color, "color"));

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
						assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo("Hello, World!", "yourValue"));
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
		/// Working With Table Row DataBands in Email Message
		/// Feature is supported by version 18.2 or greater
		/// </summary>
		public static void WorkingWithTableRowDataBandsEmail()
		{
			//ExStart:WorkingWithTableRowDataBandsEmail
			//Setting up source
			string strDocumentTemplate = "Email Templates/Working With Table Row Data Bands.msg";
			//Setting up destination 
			string strDocumentReport = "Email Reports/Working With Table Row Data Bands.msg";
			try
			{
				// Assemble a document using the external document table as a data source.
				DocumentAssembler assembler = new DocumentAssembler();
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.ExcelData(), "ds"));

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			//ExEnd:WorkingWithTableRowDataBandsEmail
		}

		/// <summary>
		/// Working With Table Row DataBands in Presentation Documents
		/// Feature is supported by version 18.2 or greater
		/// </summary>
		public static void WorkingWithTableRowDataBandsPresentation()
		{
			//ExStart:WorkingWithTableRowDataBandsPresentation
			//Setting up source 
			string strDocumentTemplate = "Presentation Templates/Working With Table Row Data Bands.pptx";
			//Setting up destination
			string strDocumentReport = "Presentation Reports/Working With Table Row Data Bands.pptx";
			try
			{
				// Assemble a document using the external document table as a data source.
				DocumentAssembler assembler = new DocumentAssembler();
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.ExcelData(), "ds"));

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			//ExEnd:WorkingWithTableRowDataBandsPresentation
		}

		/// <summary>
		/// Working With Table Row DataBands in Spread Sheet Document
		/// Feature is supported by version 18.2 or greater
		/// </summary>
		public static void WorkingWithTableRowDataBandsSpreadSheet()
		{
			//ExStart:WorkingWithTableRowDataBandsSpreadSheet
			//Setting up source 
			string strDocumentTemplate = "Spreadsheet Templates/Working With Table Row Data Bands.xlsx";
			//Setting up destination
			string strDocumentReport = "Spreadsheet Reports/Working With Table Row Data Bands.xlsx";
			try
			{
				// Assemble a document using the external document table as a data source.
				DocumentAssembler assembler = new DocumentAssembler();
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.ExcelData(), "ds"));

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			//ExEnd:WorkingWithTableRowDataBandsSpreadSheet
		}

		/// <summary>
		/// Working With Table Row DataBands in Words Processing Document
		/// Feature is supported by version 18.2 or greater
		/// </summary>
		public static void WorkingWithTableRowDataBandsWord()
		{
			//ExStart:WorkingWithTableRowDataBandsWord
			//Setting up source 
			string strDocumentTemplate = "Word Templates/Working With Table Row Data Bands.docx";
			//Setting up destination
			string strDocumentReport = "Word Reports/Working With Table Row Data Bands.docx";
			try
			{
				// Assemble a document using the external document table as a data source.
				DocumentAssembler assembler = new DocumentAssembler();
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.ExcelData(), "ds"));

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
			//Setting up source open document template
			//Setting up source email template
			const String strEmailTemplate = "Email Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.msg";
			//Setting up destination email report 
			const String strEmailReport = "Email Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.msg";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				string title = "Total Order Quantity by Quarters";
	
				var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData(),title);
				var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders","title");

				//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in email format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strEmailTemplate),
					CommonUtilities.SetDestinationDocument(strEmailReport),
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
			//Setting up source open document template
			const String strDocumentTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.xlsx";
			//Setting up destination open document report 
			const String strDocumentReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.xlsx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				string title = "Total Order Quantity by Quarters";
				//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), new DataSourceInfo(title, "title"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.pptx";
			//Setting up destination open document report 
			const String strDocumentReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.pptx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				string title = "Total Order Quantity by Quarters";
				//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), new DataSourceInfo(title, "title"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/In-Table List with Running (Progressive) Total_BackgroundColor.docx";
			//Setting up destination open document report 
			const String strDocumentReport = "Word Reports/In-Table List with Running (Progressive) Total_BackgroundColor.docx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				string color = "red";
				//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), new DataSourceInfo(color, "color"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering_dynamic_title.docx";
			//Setting up destination open document report 
			const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_dynamic_title.docx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				string title = "Total Order Quantity by Quarters";
				//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), new DataSourceInfo(title, "title"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering_RemoveIf.docx";
			//Setting up destination open document report 
			const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_RemoveIf.docx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				int mode = 2;
				//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"), new DataSourceInfo(mode, "mode"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			//ExEnd:RemoveSelectiveChartSeries
		}

		public static void GenerateBulletedList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Bulleted List_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Bulleted list report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Bulleted List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Bulleted List Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bulleted List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Bulleted List_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate bulleted list in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Bulleted List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Bulleted List Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Bulleted List_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate bulleted list in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/Bulleted List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/Bulleted List Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bulleted List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/Bulleted List.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/Bulleted List Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							// This is needed solely for images in HTML documents.
							assembler.KnownTypes.Add(typeof(GroupDocs.AssemblyExamples.BusinessLayer.CommonUtilities.FileUtil));
							//Call AssembleDocument to generate Bulleted List Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source text document template
						const String strTextTemplate = "Text Templates/Bulleted List.txt";
						//Setting up destination text document report 
						const String strTextReport = "Text Reports/Bulleted List Report.txt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Bulleted List Report in text format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strTextTemplate), CommonUtilities.SetDestinationDocument(strTextReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/Bulleted List.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/Bulleted List Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							
							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetProductsData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

							//Call AssembleDocument to generate Bulleted List Report in email format
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
						//ExEnd:GenerateBulletedListinEmailFormat
					}
					break;

			}
		}
		public static void GenerateChartWithFilteringGroupingAndOrdering(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering_Json Report.docx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Chart with Filtering, Grouping, and Ordering Report.docx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_Json Report.xlsx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx";
						//Setting up destination open document report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering Report.xlsx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_Json Report.pptx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strPresentationTemplate = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx";
						//Setting up destination open document report 
						const String strPresentationReport = "Presentation Reports/Chart with Filtering, Grouping, and Ordering Report.pptx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/Chart with Filtering, Grouping, and Ordering.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/Chart with Filtering, Grouping, and Ordering Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();

							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

							//Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in email format
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
		public static void GenerateCommonList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Common List_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Common List report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Common List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Common List Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Common List_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Common List report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Common List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Common List Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
						const String strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/Common List_XML Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Common List_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Common List_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Common List report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/Common List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/Common List Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source html template
						const String strDocumentTemplate = "HTML Templates/Common List.html";
						//Setting up destination html report 
						const String strDocumentReport = "HTML Reports/Common List Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source text document template
						const String strTxtDocumentTemplate = "Text Templates/Common List.txt";
						//Setting up destination text document report 
						const String strTxtDocumentReport = "Text Reports/Common List Report.txt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in text document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strTxtDocumentTemplate), CommonUtilities.SetDestinationDocument(strTxtDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source email template
						const String strEmailDocumentTemplate = "Email Templates/Common List.msg";
						//Setting up destination email report 
						const String strEmailDocumentReport = "Email Reports/Common List Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							
							var dataSources = DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.PopulateData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");
							//Call AssembleDocument to generate Common List Report in email document format
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
		public static void GenerateCommonMasterDetail(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Common Master-Detail_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Common Master-Detail_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Common master-detail report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Common Master-Detail_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Common Master-Detail Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common Master-Detail Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Common Master-Detail_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Common master-detail report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods";
						//Setting up destination open document report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Common Master-Detail Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Common Master-Detail_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Common Master-Detail_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Common master-detail report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open spreadsheet template
						const String strPresentationTemplate = "Presentation Templates/Common Master-Detail_OpenDocument.odp";
						//Setting up destination open document report 
						const String strPresentationReport = "Presentation Reports/Common Master-Detail Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source html template
						const String strDocumentTemplate = "HTML Templates/Common Master-Detail.html";
						//Setting up destination html report 
						const String strDocumentReport = "HTML Reports/Common Master-Detail Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							// This is needed solely for images in HTML documents.
							assembler.KnownTypes.Add(typeof(GroupDocs.AssemblyExamples.BusinessLayer.CommonUtilities.FileUtil));
							//Call AssembleDocument to generate Common Master-Detail Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source text document template
						const String strTxtDocumentTemplate = "Text Templates/Common Master-Detail.txt";
						//Setting up destination text document report 
						const String strTxtDocumentReport = "Text Reports/Common Master-Detail Report.txt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Common Master-Detail Report in text document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strTxtDocumentTemplate), CommonUtilities.SetDestinationDocument(strTxtDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source email document template
						const String strEmailDocumentTemplate = "Email Templates/Common Master-Detail.msg";
						//Setting up destination email document report 
						const String strEmailDocumentReport = "Email Reports/Common Master-Detail Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							
							var dataSources = DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.PopulateData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

							//Call AssembleDocument to generate Common Master-Detail Report in email document format
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
		public static void GenerateInParagraphList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/In-Paragraph List_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Paragraph List report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/In-Paragraph List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/In-Paragraph List Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Paragraph List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/In-Paragraph List_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Paragraph List report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Paragraph List Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/In-Paragraph List_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/In-Paragraph List_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Paragraph List report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/In-Paragraph List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Paragraph List Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Paragraph List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source html template
						const String strDocumentTemplate = "HTML Templates/In-Paragraph List.html";
						//Setting up destination html report 
						const String strDocumentReport = "HTML Reports/In-Paragraph List Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Paragraph List Report in htmlformat
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source text document template
						const String strDocumentTemplate = "Text Templates/In-Paragraph List.txt";
						//Setting up destination text document report 
						const String strDocumentReport = "Text Reports/In-Paragraph List Report.txt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Paragraph List Report in text document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source email document template
						const String strEmailDocumentTemplate = "Email Templates/In-Paragraph List.msg";
						//Setting up destination email document report 
						const String strEmailDocumentReport = "Email Reports/In-Paragraph List Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();

							var dataSources = DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.GetProductsData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

							//Call AssembleDocument to generate In-Paragraph List Report in email document format
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
		public static void GenerateInTableListWithAlternateContent(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/In-Table List with Alternate Content_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Alternate Content report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/In-Table List with Alternate Content_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/In-Table List with Alternate Content Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/In-Table List with Alternate Content_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Alternate Content report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Alternate Content Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/In-Table List with Alternate Content_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Alternate Content report in presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Table List with Alternate Content Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/In-Table List with Alternate Content.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/In-Table List with Alternate Content Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Alternate Content Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/In-Table List with Alternate Content.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/In-Table List with Alternate Content Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							
							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

							//Call AssembleDocument to generate In-Table List with Alternate Content Report in email format
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
		public static void GenerateInTableListWithFilteringGroupingAndOrdering(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/In-Table List with Filtering, Grouping, and Ordering Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/In-Table List with Filtering, Grouping, and Ordering.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/In-Table List with Filtering, Grouping, and Ordering Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/In-Table List with Filtering, Grouping, and Ordering.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/In-Table List with Filtering, Grouping, and Ordering Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							
							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

							//Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in email format
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
		public static void GenerateInTableListWithHighlightedRows(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Highlighted Rows report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/In-Table List with Highlighted Rows Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Highlighted Rows report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Highlighted Rows Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersDataDB(), "orders"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List with Highlighted Rows report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerOrderDataFromJson(), "orders"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Table List with Highlighted Rows Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/In-Table List with Highlighted Rows.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/In-Table List with Highlighted Rows Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Highlighted Rows Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/In-Table List with Highlighted Rows.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/In-Table List with Highlighted Rows Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();

							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.GetOrdersData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

							//Call AssembleDocument to generate In-Table List with Highlighted Rows Report in email format
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
		public static void GenerateInTableList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/In-Table List_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/In-Table List_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/In-Table List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/In-Table List Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/In-Table List_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/In-Table List_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/In-Table List_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/In-Table List_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table List report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/In-Table List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Table List Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/In-Table List.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/In-Table List Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/In-Table List.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/In-Table List Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();

							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.PopulateData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

							//Call AssembleDocument to generate In-Table List Report in email format
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
		public static void GenerateInTableMasterDetail(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/In-Table Master-Detail_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/In-Table Master-Detail_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table Master-Detail report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/In-Table Master-Detail_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/In-Table Master-Detail Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table Master-Detail Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/In-Table Master-Detail_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table Master-Detail report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Table Master-Detail Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersDataDB(), "customers"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomersAndOrdersDataDT(), "ds"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/In-Table Master-Detail_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/In-Table Master-Detail_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate In-Table Master-Detail report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerDataFromJson(), "customers"));
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
						//Setting up source open spreadsheet template
						const String strPresentationTemplate = "Presentation Templates/In-Table Master-Detail_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Table Master-Detail Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/In-Table Master-Detail.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/In-Table Master-Detail Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table Master-Detail Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.PopulateData(), "customers"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/In-Table Master-Detail.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/In-Table Master-Detail Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();

							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.PopulateData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

							//Call AssembleDocument to generate In-Table Master-Detail Report in email format
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

		public static void GenerateInTableListWithProgressiveTotal(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
		{
			switch (strDocumentFormat)
			{
				case "document":
					//ExStart:GenerateInTableListWithProgressiveTotalinOpenDocumentProcessingFormat
					//Setting up source open document template
					const String strDocumentTemplate = "Word Templates/In-Table List with Running (Progressive) Total.docx";
					//Setting up destination open document report 
					const String strDocumentReport = "Word Reports/In-Table List with Running (Progressive) Total Report.docx";
					try
					{
						var testData = DataLayer.GetOrdersData();
						//Instantiate DocumentAssembler class
						DocumentAssembler assembler = new DocumentAssembler();
						//Call AssembleDocument to generate In-Table List with Progressive(Running) total Report in open document format
						assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/In-Table List with Running (Progressive) Total.xlsx";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/In-Table List with Running (Progressive) Total Report.xlsx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Progressive(Running) total Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/In-Table List with Running (Progressive) Total.pptx";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/In-Table List with Running (Progressive) Total Report.pptx";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Progressive(running) total Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source html template
						const String strHtmlTemplate = "HTML Templates/In-Table List with Running (Progressive) Total.html";
						//Setting up destination html report 
						const String strHtmlReport = "HTML Reports/In-Table List with Running (Progressive) Total.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate In-Table List with Progressive(Running) Total Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
						//Setting up source email template
						const String strEmailTemplate = "Email Templates/In-Table List with Running (Progressive) Total.msg";
						//Setting up destination email report 
						const String strEmailReport = "Email Reports/In-Table List with Running (Progressive) Total.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();

							var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate,DataLayer.GetOrdersData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "orders");

							//Call AssembleDocument to generate In-Table List with Progressive(Running) Total Report in email format
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
		public static void GenerateMulticoloredNumberedList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
		{
			switch (strDocumentFormat)
			{
				case "document":
					if (isDatabase)
					{
						//ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenDocumentProcessingDocument
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Multicolored Numbered List_DB Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Multicolored Numbered List_DT Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_XML_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Multicolored Numbered List_XML Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Multicolored Numbered List_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Multicolored Numbered List report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Multicolored Numbered List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Multicolored Numbered List Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List_DB Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List_DT Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_XML_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List_XML Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Multicolored Numbered List_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Multicolored Numbered List report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Multicolored Numbered List Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/Multicolored Numbered List_DB Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Multicolored Numbered List_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Multicolored Numbered List report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/Multicolored Numbered List Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source html template
						const String strDocumentTemplate = "HTML Templates/Multicolored Numbered List.html";
						//Setting up destination html report 
						const String strDocumentReport = "HTML Reports/Multicolored Numbered List Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Multicolored Numbered List Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source email template
						const String strEmailDocumentTemplate = "Email Templates/Multicolored Numbered List.msg";
						//Setting up destination email report 
						const String strEmailDocumentReport = "Email Reports/Multicolored Numbered List Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							var dataSources = DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.GetProductsData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

							//Call AssembleDocument to generate Multicolored Numbered List Report in html format
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
					}
					break;

			}
		}
		public static void GenerateNumberedList(string strDocumentFormat, bool isDatabase, bool isDataSet, bool isDataSourceXML, bool isJson)
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
						//Setting up destination 
						const String strDocumentReport = "Word Reports/Numbered List_OpenDocument_Json Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Numbered List report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open document template
						const String strDocumentTemplate = "Word Templates/Numbered List_OpenDocument.odt";
						//Setting up destination open document report 
						const String strDocumentReport = "Word Reports/Numbered List Report.odt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Numbered List Report in open document format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
						//Setting up destination 
						const String strDocumentReport = "Spreadsheet Reports/Numbered List_OpenDocument_Json Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Numbered List report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open spreadsheet template
						const String strSpreadsheetTemplate = "Spreadsheet Templates/Numbered List_OpenDocument.ods";
						//Setting up destination open spreadsheet report 
						const String strSpreadsheetReport = "Spreadsheet Reports/Numbered List Report.ods";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Numbered List Report in open spreadsheet format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDataDB(), "products"));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsDT()));
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
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
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
						//setting up source 
						const String strDocumentTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
						//Setting up destination 
						const String strDocumentReport = "Presentation Reports/Numbered List_OpenDocument_Json Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
																				  //Call AssembleDocument to generate Numbered List report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
						//Setting up source open presentation template
						const String strPresentationTemplate = "Presentation Templates/Numbered List_OpenDocument.odp";
						//Setting up destination open presentation report 
						const String strPresentationReport = "Presentation Reports/Numbered List Report.odp";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Numbered List Report in open presentation format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source html template
						const String strDocumentTemplate = "HTML Templates/Numbered List.html";
						//Setting up destination html report 
						const String strDocumentReport = "HTML Reports/Numbered List Report.html";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Numbered List Report in html format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source text document template
						const String strDocumentTemplate = "Text Templates/Numbered List.txt";
						//Setting up destination text document report 
						const String strDocumentReport = "Text Reports/Numbered List Report.txt";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
							//Call AssembleDocument to generate Numbered List Report in text document  format
							assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetProductsData(), "products"));
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
						//Setting up source email template
						const String strEmailDocumentTemplate = "Email Templates/Numbered List.msg";
						//Setting up destination email report 
						const String strEmailDocumentReport = "Email Reports/Numbered List Report.msg";
						try
						{
							//Instantiate DocumentAssembler class
							DocumentAssembler assembler = new DocumentAssembler();
				
							var dataSources = DataLayer.EmailDataSourceObject(strEmailDocumentTemplate, DataLayer.GetProductsData());
							var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "products");

							//Call AssembleDocument to generate Numbered List Report in email format
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
            //setting up source 
            const String strDocumentTemplate = "Word Templates/Numbered List_RestartNum.docx";
            //Setting up destination 
            const String strDocumentReport = "Word Reports/Numbered List_RestartNum.docx";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
                                                                      //Call AssembleDocument to generate Numbered List report in open document format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
            //setting up source 
            const String strDocumentTemplate = "Email Templates/Numbered List_RestartNum.msg";
            //Setting up destination 
            const String strDocumentReport = "Email Reports/Numbered List_RestartNum.msg";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();//initialize object of DocumentAssembler class 
                                                                      //Call AssembleDocument to generate Numbered List report in open document format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/Lazy And Recursive.docx";
			//Setting up destination open document report 
			const String strDocumentReport = "Word Reports/Lazy And Recursive Report.docx";
			try
			{
				//Instantiate DynamicEntity class
				DynamicEntity dentity = new DynamicEntity(Guid.NewGuid());
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to generate Single Row Report in open document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(dentity, "root"));
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
				//Setting up source open document template
				const String strDocumentTemplate = "Word Templates/Multiple DS.odt";
				//Setting up destination open document report 
				const String strDocumentReport = "Word Reports/Multiple DS.odt";
				try
				{
					//Instantiate DocumentAssembler class
					DocumentAssembler assembler = new DocumentAssembler();
					//Call AssembleDocument to generate report
					assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
				//Setting up source open document template
				const String strDocumentTemplate = "Spreadsheet Templates/Multiple DS.ods";
				//Setting up destination open document report 
				const String strDocumentReport = "Spreadsheet Reports/Multiple DS.ods";
				try
				{
					//Instantiate DocumentAssembler class
					DocumentAssembler assembler = new DocumentAssembler();
					//Call AssembleDocument to generate report
					assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
				//Setting up source open document template
				const String strDocumentTemplate = "Presentation Templates/Multiple DS.odp";
				//Setting up destination open document report 
				const String strDocumentReport = "Presentation Reports/Multiple DS.odp";
				try
				{
					//Instantiate DocumentAssembler class
					DocumentAssembler assembler = new DocumentAssembler();
					//Create an array of data source objects
					object[] dataSourceObj = new object[] { DataLayer.GetAllDataFromXML(), DataLayer.GetProductsDataJson() };
					//Create an array of data source string
					string[] dataSourceString = new string[] { "ds", "products" };

					//Call AssembleDocument to generate report
					assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"), new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/String_Numeric_Formatting.odt";
			//Setting up destination open document report 
			const String strDocumentReport = "Word Reports/String_Numeric_Formatting Report.odt";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to generate   Report in open document format
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
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
			//Setting up source open document template
			const String strDocumentTemplate = "Word Templates/OuterDocInsertion.odt";
			//Setting up destination open document report 
			const String strDocumentReport = "Word Reports/OuterDocInsertion Report.odt";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to generate  Report in open document format
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

		/// <summary>
		/// Generate barcodes
		/// </summary>
		/// <param name="strDocumentFormat"></param>
		public static void AddBarCodes(string strDocumentFormat)
		{
			switch (strDocumentFormat)
			{
				case "document":
					//ExStart:AddBarCodesDocumentProcessingFormat
					//Setting up source open document template
					const String strDocumentTemplate = "Word Templates/Barcode.docx";
					//Setting up destination open document report 
					const String strDocumentReport = "Word Reports/Barcode.docx";
					try
					{
						//Instantiate DocumentAssembler class
						DocumentAssembler assembler = new DocumentAssembler();
						//Call AssembleDocument to generate   Report in open document format
						assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message);
					}
					//ExEnd:AddBarCodesDocumentProcessingFormat 
					break;

				case "spreadsheet":
					//ExStart:AddBarCodesSpreadsheet
					//Setting up source open spreadsheet template
					const String strSpreadsheetTemplate = "Spreadsheet Templates/Barcode.xlsx";
					//Setting up destination open spreadsheet report 
					const String strSpreadsheetReport = "Spreadsheet Reports/Barcode.xlsx";
					try
					{
						//Instantiate DocumentAssembler class
						DocumentAssembler assembler = new DocumentAssembler();
						//Call AssembleDocument to generate  Report in open spreadsheet format
						assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message);
					}
					//ExEnd:AddBarCodesSpreadsheet 
					break;

				case "presentation":
					//ExStart:AddBarCodesPowerPoint
					//Setting up source open presentation template
					const String strPresentationTemplate = "Presentation Templates/Barcode.pptx";
					//Setting up destination open presentation report 
					const String strPresentationReport = "Presentation Reports/Barcode.pptx";
					try
					{
						//Instantiate DocumentAssembler class
						DocumentAssembler assembler = new DocumentAssembler();
						//Call AssembleDocument to generate  Report in open presentation format
						assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message);
					}
					//ExEnd:AddBarCodesPowerPoint 
					break;
			}
		}

		/// <summary>
		/// Update word processing documents fields while assembling 
		/// </summary>
		/// <param name="documentType"></param>
		public static void UpdateWordDocFields(string documentType)
		{

			if (documentType == "document")
			{
				//ExStart:UpdateWordDocFields
				//Setting up source document template
				const String strDocumentTemplate = "Word Templates/Update_Field_XML.docx";
				//Setting up destination document report 
				const String strDocumentReport = "Word Reports/Update_Field_XML Report.docx";
				DocumentAssembler assembler = new DocumentAssembler();
				assembler.Options |= DocumentAssemblyOptions.UpdateFieldsAndFormulas;
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
				//ExEnd:UpdateWordDocFields
			}
			else if (documentType == "spreadsheet")
			{
				//ExStart:updateformula
				//Setting up source document template
				const String strDocumentTemplate = "Spreadsheet Templates/Update-Fomula.xlsx";
				//Setting up destination document report 
				const String strDocumentReport = "Spreadsheet Reports/Update-Fomula Report.xlsx";
				DocumentAssembler assembler = new DocumentAssembler();
				assembler.Options |= DocumentAssemblyOptions.UpdateFieldsAndFormulas;
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
				//ExEnd:updateformula
			}

		}

		/// <summary>
		/// Support for Analogue of Microsoft Word NEXT Field
		/// </summary>
		public static void NextIteration()
		{
			//ExStart:nextiteration
			//Setting up source document template
			const String strDocumentTemplate = "Word Templates/Using Next.docx";
			//Setting up destination document report
			const String strDocumentReport = "Word Reports/Using Next Report.docx";
			try
			{
				//Instantiate DocumentAssembler class
				DocumentAssembler assembler = new DocumentAssembler();
				//Call AssembleDocument to generate Report 
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.GetAllDataFromXML(), "ds"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			//ExEnd:nextiteration
		}

		/// <summary>
		/// Generate report from excel data source
		/// </summary>
		public static void UseSpreadsheetAsDataSource()
		{
			//ExStart:UseSpreadsheetAsDataSource
			string strDocumentTemplate = "Word Templates/Using Spreadsheet as Table of Data.docx";
			string strDocumentReport = "Word Reports/Using Spreadsheet as Table of Data_Output.docx";

			// Assemble a document using the external document table as a data source.
			DocumentAssembler assembler = new DocumentAssembler();
			assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.ExcelData(), "contracts"));
			//ExEnd:UseSpreadsheetAsDataSource
		}

		/// <summary>
		/// Generate report from presentation data source
		/// </summary>
		public static void UsePresentationTableAsDataSource()
		{
			string strDocumentTemplate = "Presentation Templates/Using Presentation as Table of Data.pptx";
			string strDocumentReport = "Presentation Reports/Using Presentation as Table of Data_Output.pptx";


			// Assemble a document using the external document table as a data source.
			DocumentAssembler assembler = new DocumentAssembler();
			assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.PresentationData(), "table"));
		}
		/// <summary>
		/// Import word proccessing table into presentation
		/// </summary>
		public static void ImportingWordProcessingTableIntoPresentation()
		{
			string strDocumentTemplate = "Presentation Templates/Importing Word Processing Table into Presentation.pptx";
			string strDocumentReport = "Presentation Reports/Importing Word Processing Table into Presentation_Output.pptx";


			// Assemble a document using the external document table as a data source.
			DocumentAssembler assembler = new DocumentAssembler();
			assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(DataLayer.ImportingWordDocToPresentation(), "table"));
		}

		/// <summary>
		/// Import Spreadsheet into HTML Document
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
				assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strHtmlTemplate), CommonUtilities.SetDestinationDocument(strHtmlReport), new DataSourceInfo(DataLayer.ImportingSpreadsheetToHtml(), "table"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			//ExEnd:ImportingSpreadsheetIntoHtmlDocument
		}

        /// <summary>
		/// Table Cell Merging in Word Processing Documents
		/// Features is supported by version 19.1 or greater
		/// </summary>
		public static void TableCellsMergingInWordProcessing()
        {
            //ExStart:TableCellsMergingInWordProcessing
            //Setting up source document template
            const String strDocumentTemplate = "Word Templates/Merging Cells Dynamically.docx";
            //Setting up destination PDF report 
            const String strDocumentReport = "PDF Reports/Merging Cells Dynamically Report.pdf";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to Merging Cells Dynamically Report in PDF format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf), new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInWordProcessing
        }
        /// <summary>
		/// Table Cell Merging in Presentations
		/// Features is supported by version 19.1 or greater
		/// </summary>
        public static void TableCellsMergingInPresentations()
        {
            //ExStart:TableCellsMergingInPresentations
            //Setting up source presentation template
            const String strDocumentTemplate = "Presentation Templates/Merging Cells Dynamically.pptx";
            //Setting up destination PDF report 
            const String strDocumentReport = "PDF Reports/Merging Cells Dynamically Report.pdf";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to Merging Cells Dynamically Report in PDF format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf), new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInPresentations
        }
        /// <summary>
		/// Table Cell Merging in Spreadsheets
		/// Features is supported by version 19.1 or greater
		/// </summary>
        public static void TableCellsMergingInSpreadsheets()
        {
            //ExStart:TableCellsMergingInSpreadsheets
            //Setting up source spreadsheet template
            const String strDocumentTemplate = "Spreadsheet Templates/Merging Cells Dynamically.xlsx";
            //Setting up destination PDF report 
            const String strDocumentReport = "PDF Reports/Merging Cells Dynamically Report.pdf";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to Merging Cells Dynamically Report in PDF format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf), new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:TableCellsMergingInSpreadsheets
        }
        /// <summary>
		/// Table Cell Merging in Emails
		/// Features is supported by version 19.1 or greater
		/// </summary>
        public static void TableCellsMergingInEmails()
        {
            //ExStart:TableCellsMergingInEmails
            //Setting up source email template
            const String strEmailTemplate = "Email Templates/Merging Cells Dynamically.msg";
            //Setting up destination email report 
            const String strEmailReport = "Email Reports/Merging Cells Dynamically Report.msg";
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                var dataSources = DataLayer.EmailDataSourceObject(strEmailTemplate, DataLayer.PopulateData());
                var dataSourcesNames = DataLayer.EmailDataSourceName(".msg", "customers");

                //Call AssembleDocument to generate Merging Cells Dynamically Report in email format
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
        /// Features is supported by version 19.3 or greater
        /// </summary>
        public static void DemoInLineSyntaxError()
        {
            //ExStart:DemoInLineSyntaxError_19.3
            //Setting up source document template
            const String strDocumentTemplate = "Word Templates/Inline Error Demo.docx";
            //Setting up destination PDF report 
            const String strDocumentReport = "PDF Reports/Inline Error Demo.pdf";
            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();

                //Enable the In-line error messaging
                assembler.Options |= DocumentAssemblyOptions.InlineErrorMessages;

                //Call AssembleDocument to show the demo Report and save the report in PDF format
                //The AssembleDocument will return a boolean value to indicate the success or failed with inline error.
                if (assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new LoadSaveOptions(FileFormat.Pdf), new DataSourceInfo(DataLayer.GetCustomerData(), "customer")))
                {
                    Console.WriteLine("No error found in template");
                }else
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
        /// Loading of template documents from HTML with resources
        /// Features is supported by version 19.5 or greater
        /// </summary>
        public static void LoadDocFromHTMLWithResource()
        {
            //ExStart:LoadDocFromHTMLWithResource_19.5

            //Setting up resource directory for input template and output
            String strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceLoad/");
            
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                
                //Assemble the document using relative directory path and file names.
                assembler.AssembleDocument(strDirectoryPath + "TestWordsResourceLoad.htm", strDirectoryPath + "TestWordsResourceLoad_Out.docx", new DataSourceInfo("It should be a jeep image.", "value"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:LoadDocFromHTMLWithResource_19.5
        }
        /// <summary>
        /// Loading of template documents from HTML with resources from an explicitly specified folder
        /// Features is supported by version 19.5 or greater
        /// </summary>
        public static void LoadDocFromHTMLWithResource_ExplicitFolder()
        {
            //ExStart:LoadDocFromHTMLWithResource_ExplicitFolder_19.5

            //Setting up resource directory for input template and output
            String strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceLoad/");

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                //Initialize LoadSaveOptions
                LoadSaveOptions loadSaveOptions = new LoadSaveOptions();
                
                //Resolve URIs from the specified alternative folder.
                loadSaveOptions.ResourceLoadBaseUri = strDirectoryPath + "Alternative";

                assembler.AssembleDocument(strDirectoryPath + "TestWordsResourceLoad.htm", strDirectoryPath + "TestWordsResourceLoad_Out.docx", loadSaveOptions, new DataSourceInfo("It should be a sport car image.", "value"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:LoadDocFromHTMLWithResource_ExplicitFolder_19.5
        }
        /// <summary>
        /// Saving of external resource files at relative path while an assembled document loaded from a non-HTML format is being saved to HTML
        /// Features is supported by version 19.5 or greater
        /// </summary>
        public static void SaveDocToHTMLWithResource()
        {
            //ExStart:SaveDocToHTMLWithResource_19.5

            //Setting up resource directory for input template and output
            String strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceSave/");
            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                //Assemble the document using relative directory path and file names.
                assembler.AssembleDocument(strDirectoryPath + "TestCellsResourceSaveOneSheet.xlsx", strDirectoryPath + "TestCellsResourceSaveOneSheet_Out.htm", new DataSourceInfo("Hello!", "value"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveDocToHTMLWithResource_19.5
        }
        /// <summary>
        /// Saving of external resource files in a specified folder at relative path while saving output to HTML
        /// Features is supported by version 19.5 or greater
        /// </summary>
        public static void SaveDocToHTMLWithResource_ExplicitFolder()
        {
            //ExStart:SaveDocToHTMLWithResource_ExplicitFolder_19.5

            //Setting up resource directory for input template and output
            String strDirectoryPath = CommonUtilities.GetSourceFolder("ResourceSave/");

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();

                //Initialize LoadSaveOptions
                LoadSaveOptions loadSaveOptions = new LoadSaveOptions();

                //Resolve URIs from the specified alternative folder.
                loadSaveOptions.ResourceSaveFolder = strDirectoryPath + "Alternative";

                assembler.AssembleDocument(strDirectoryPath + "TestWordsResourceSave.docx", strDirectoryPath + "TestWordsResourceSave_Out.htm", loadSaveOptions, new DataSourceInfo("Hello!", "value"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveDocToHTMLWithResource_ExplicitFolder_19.5
        }
        /// <summary>
        /// Saving an assembled Markdown document to a Word Processing format using file extension.
        /// Features is supported by version 19.8 or greater
        /// </summary>
        public static void SaveMdtoWord_UsingExtension(String DocumentName)
        {
            //ExStart:SaveMdtoWord_UsingExtension_19.8

            //Setting up source document template
            String strDocumentTemplate = "Markdown Templates/"+ DocumentName;
            //Setting up destination Markdown reports 
            String strDocumentReport = "Word Reports/"+ DocumentName + " Out.docx";
            //Setting up description variable
            const string description = "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
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
        /// Features is supported by version 19.8 or greater
        /// </summary>
        public static void SaveWordOrEmailtoMD_UsingExtension()
        {
            //ExStart:SaveWordtoMD_UsingExtension_19.8

            //Setting up source document template (Email or Word Document)
            const String strDocumentTemplate = "Word Templates/ReadMe.docx";

            //Setting up destination Markdown reports 
            const String strDocumentReport = "Markdown Reports/ReadMe Out.md";
            //Setting up description variable
            const string description = "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
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
        /// Features is supported by version 19.8 or greater
        /// </summary>
        public static void SaveWordOrEmailtoMD_Explicit()
        {
            //ExStart:SaveWordtoMD_Explicit_19.8
           
            try
            {
                //Setting up source document template (Email or Word Document)
                Stream sourceStream = File.OpenRead(CommonUtilities.GetSourceDocument("Word Templates/ReadMe.docx"));
                //Setting up destination Markdown Reports 
                FileStream targetStream = File.Create(CommonUtilities.SetDestinationDocument("Markdown Reports/ReadMe Out.md"));

                //Setting up description variable
                const string description = "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
                                              "office and email file formats based upon template documents and data obtained from various sources " +
                                              "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

                DataSourceInfo dataSourceInfo1 = new DataSourceInfo("GroupDocs.Assembly for .NET(Using Stream)", "product");
                DataSourceInfo dataSourceInfo2 = new DataSourceInfo(description, "description");

                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(sourceStream, targetStream, new LoadSaveOptions(FileFormat.Markdown), dataSourceInfo1,dataSourceInfo2);
               
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SaveWordtoMD_UsingExtension_19.8
        }
        /// <summary>
        /// Working with JSON data sources
        /// Features is supported by version 19.10 or greater
        /// </summary>
        public static void SimpleJsonDS_Demo()
        {
            //ExStart:SimpleJsonDS_Demo_19.10

            try
            {
                //Setting up source document template (Email or Word Document)
                const String strDocumentTemplate = "Word Templates/SimpleDatasetDemo.docx";

                //Setting up destination for reports 
                const String strDocumentReport = "Word Reports/SimpleJsonDSDemo Out.docx";

                //Setting up destination Markdown reports 
                const String strDataSource = "JSON DataSource/ManagerData.json";


                //initialize data source
                JsonDataSource dataSource = new JsonDataSource(CommonUtilities.GetDataSourceDocument(strDataSource));

            DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(dataSource, "managers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SimpleJsonDS_Demo_19.10
        }
        /// <summary>
        /// Working with XML data sources
        /// Features is supported by version 19.10 or greater
        /// </summary>
        public static void SimpleXMLDS_Demo()
        {
            //ExStart:SimpleXMLDS_Demo_19.10

            try
            {
                //Setting up source document template (Email or Word Document)
                const String strDocumentTemplate = "Word Templates/SimpleDatasetDemo.docx";

                //Setting up destination for reports 
                const String strDocumentReport = "Word Reports/SimpleXMLDSDemo Out.docx";

                //Setting up Data Source file
                const String strDataSource = "XML DataSource/Managers.xml";


                //initialize data source
                XmlDataSource dataSource = new XmlDataSource(CommonUtilities.GetDataSourceDocument(strDataSource));

                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(dataSource, "managers"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SimpleXMLDS_Demo_19.10
        }
        /// <summary>
        /// Working with csv data sources
        /// Features is supported by version 19.10 or greater
        /// </summary>
        public static void SimpleCsvDS_Demo()
        {
            //ExStart:SimpleCsvDS_Demo_19.10

            try
            {
                //Setting up source document template (Email or Word Document)
                const String strDocumentTemplate = "Text Templates/CsvDatasetDemo.txt";

                //Setting up destination for reports 
                const String strDocumentReport = "Word Reports/SimpleCsvDSDemo Out.docx";

                //Setting up destination Markdown reports 
                const String strDataSource = "Excel DataSource/Person.csv";

                //Setting up CSV data load options
                CsvDataLoadOptions options = new CsvDataLoadOptions(true);
             
                //initialize data source
                CsvDataSource dataSource = new CsvDataSource(CommonUtilities.GetDataSourceDocument(strDataSource), options);

                //Assemble the document
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
        /// Insert Image Dynamically in Word Document
        /// Features is supported by version 20.3 or greater
        /// </summary>
        public static void InsertImageDynamicallyInWord()
        {
            //ExStart:InsertImageDynamicallyInWord_20.3

            try
            {
                //Setting up source document template (Email or Word Document)
                const String strDocumentTemplate = "Word Templates/DynamicImageDemo.docx";

                //Setting up destination for reports 
                const String strDocumentReport = "Word Reports/DynamicImageDemo Out.docx";

                //Assemble the document
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), 
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(CommonUtilities.GetImageFolder("no-photo.jpg"),"image_expression")); 
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
            //ExEnd:InsertImageDynamicallyInWord_20.3
        }

        /// <summary>
        /// Insert Document Dynamically in Word Document
        /// Features is supported by version 20.3 or greater
        /// </summary>
        public static void InsertDocumentDynamicallyInWord()
        {
            //ExStart:InsertDocumentDynamicallyInWord_20.3

            try
            {
                //Setting up source document template (Email or Word Document)
                const String strDocumentTemplate = "Word Templates/DynamicDocInsert.docx";

                //Setting up destination for reports 
                const String strDocumentReport = "Word Reports/DynamicDocInsert Out.docx";

                //Assemble the document
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(CommonUtilities.GetOuterDocumentFolder("OuterDocument.docx"), "document_expression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:InsertDocumentDynamicallyInWord_20.3
        }


        /// <summary>
        /// Set checkbox value dynamically in Word document
        /// Features is supported by version 20.3 or greater
        /// </summary>
        public static void SetCheckboxValueDynamicallyInWord(Boolean boolVal)
        {
            //ExStart:SetCheckboxValueDynamicallyInWord_20.3

            try
            {
                //Setting up source document template (Email or Word Document)
                const String strDocumentTemplate = "Word Templates/CheckBoxValueSetDemo.docx";

                //Setting up destination for reports 
                const String strDocumentReport = "Word Reports/CheckBoxValueSetDemo Out.docx";

                //Assemble the document
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                    CommonUtilities.SetDestinationDocument(strDocumentReport), new DataSourceInfo(boolVal, "conditional_expression"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:SetCheckboxValueDynamicallyInWord_20.3
        }

    }
    
}

