
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports GroupDocs.Assembly
Imports GroupDocs.AssemblyExamples.BusinessLayer
Imports GroupDocs.AssemblyExamples.BusinessLayer.GroupDocs.AssemblyExamples.BusinessLayer



Namespace GroupDocs.AssemblyExamples
    Public NotInheritable Class GenerateReport
        Private Sub New()
        End Sub




        Public Shared Sub GenerateBubbleChart(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateBubbleChartFromDatabaseinOpenDocumentProcessingFormat
                        'setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateBubbleChartFromDataSetinOpenDocumentProcessingFormat
                        'setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart_DT Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateBubbleChartFromXMLinOpenDocumentProcessingFormat
                        'setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart_XML Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromXMLinOpenDocumentProcessingFormat

                        End Try
                    Else
                        'ExStart:GenerateBubbleChartinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateBubbleChartFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bubble Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bubble Chart_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateBubbleChartFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bubble Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bubble Chart_DT Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateBubbleChartFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bubble Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bubble Chart_XML Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateBubbleChartinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bubble Chart.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bubble Chart Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateBubbleChartFromDatabaseinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bubble Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bubble Chart_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateBubbleChartFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bubble Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bubble Chart_DT Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateBubbleChartFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bubble Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bubble Chart_XML Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateBubbleChartinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bubble Chart.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bubble Chart Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateBulletedList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateBulletedListFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateBulletedListFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateBulletedListFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateBulletedListinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateBulletedListFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bulleted List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bulleted List_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateBulletedListFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bulleted List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bulleted List_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateBulletedListFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bulleted List_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bulleted List_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateBulletedListinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bulleted List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bulleted List Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateBulletedListFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bulleted List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bulleted List_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateBulletedListFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bulleted List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bulleted List_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateBulletedListFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bulleted List_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bulleted List_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateBulletedListinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bulleted List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bulleted List Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateChartWithFilteringGroupingAndOrdering(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering_DT Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering_XML.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering_XML Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DT Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_XML.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_XML Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DT Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_XML.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_XML Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateCommonList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateCommonListFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common List_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateCommonListFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common List_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateCommonListFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common List_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonListinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common List Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateCommonListFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common List_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateCommonListFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common List_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateCommonListFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common List_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonListinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common List Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateCommonListFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common List_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateCommonListFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common List_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateCommonListFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common List_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonListinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common List Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateCommonMasterDetail(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateCommonMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateCommonMasterDetailFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateCommonMasterDetailFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonMasterDetailinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateCommonMasterDetailFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common Master-Detail_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateCommonMasterDetailFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common Master-Detail_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateCommonMasterDetailFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common Master-Detail_DB_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common Master-Detail_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonMasterDetailinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common Master-Detail Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateCommonMasterDetailFromDatabaseinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common Master-Detail_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateCommonMasterDetailFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common Master-Detail_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateCommonMasterDetailFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common Master-Detail_DB_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common Master-Detail_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonMasterDetailinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common Master-Detail_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common Master-Detail Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInParagraphList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInParagraphListFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInParagraphListFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInParagraphListFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateInParagraphListinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInParagraphListFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Paragraph List_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInParagraphListFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Paragraph List_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInParagraphListFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Paragraph List_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Paragraph List_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInParagraphListinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Paragraph List Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInParagraphListFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Paragraph List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Paragraph List_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInParagraphListFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Paragraph List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Paragraph List_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInParagraphListFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Paragraph List_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Paragraph List_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInParagraphListinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Paragraph List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Paragraph List Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableListWithAlternateContent(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content_DT_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithAlternateContentFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithAlternateContentinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content_DB_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content_DT_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithAlternateContentFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithAlternateContentinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content_DB_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Alternate Content_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content_DT_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Alternate Content_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithAlternateContentFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Alternate Content_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithAlternateContentinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Alternate Content Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableListWithFilteringGroupingAndOrdering(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginOpenDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableListWithHighlightedRows(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows_DT_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromXMLinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinOpenDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsinOpenDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows_DB_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenSpreadsheetDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows_DT_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinOpenSpreadsheetDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromXMLinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinOpenSpreadsheetDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsinOpenSpreadsheetDocument
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows_DB_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows_DT_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListFromDatabaseinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDatabaseinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListFromDataSetinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List_DT_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDataSetinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListFromXMLinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromXMLinOpenDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListinOpenDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List_DB_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List_DT_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List_DB_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List_DT_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableMasterDetail(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail_DB_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableMasterDetailFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail_DT_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableMasterDetailFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableMasterDetailinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableMasterDetailFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail_DB_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table Master-Detail_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableMasterDetailFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail_DT_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table Master-Detail_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableMasterDetailFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table Master-Detail_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableMasterDetailinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table Master-Detail Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableMasterDetailFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table Master-Detail_DB_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table Master-Detail_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableMasterDetailFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table Master-Detail_DT_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table Master-Detail_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableMasterDetailFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table Master-Detail_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table Master-Detail_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableMasterDetailinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table Master-Detail_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table Master-Detail Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateMulticoloredNumberedList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDatabaseinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateMulticoloredNumberedListFromDataSetinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDataSetinOpenDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateMulticoloredNumberedListFromXMLinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateMulticoloredNumberedListinOpenDocumentProcessingDocument
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListinOpenDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Multicolored Numbered List_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDatabaseinOpenSpreadsheetDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateMulticoloredNumberedListFromDataSetinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Multicolored Numbered List_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDataSetinOpenSpreadsheetDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateMulticoloredNumberedListFromXMLinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Multicolored Numbered List_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenSpreadsheetDocument
                        End Try
                    Else
                        'ExStart:GenerateMulticoloredNumberedListinOpenSpreadsheetDocument
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Multicolored Numbered List Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListinOpenSpreadsheetDocument
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateMulticoloredNumberedListFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Multicolored Numbered List_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateMulticoloredNumberedListFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Multicolored Numbered List_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateMulticoloredNumberedListFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Multicolored Numbered List_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Multicolored Numbered List_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateMulticoloredNumberedListinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Multicolored Numbered List Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateNumberedList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateNumberedListFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateNumberedListFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDataSetinOpenDocumentProcessingFormat 
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateNumberedListFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List_XML_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromXMLinOpenDocumentProcessingFormat 
                        End Try
                    Else
                        'ExStart:GenerateNumberedListinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateNumberedListFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Numbered List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Numbered List_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateNumberedListFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Numbered List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Numbered List_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateNumberedListFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Numbered List_XML_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Numbered List_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateNumberedListinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Numbered List_OpenDocument.ods"
                        'Setting up destination open spreadsheet report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Numbered List Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateNumberedListFromDatabaseinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Numbered List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Numbered List_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateNumberedListFromDataSetinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Numbered List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Numbered List_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateNumberedListFromXMLinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Numbered List_XML_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Numbered List_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateNumberedListinOpenPresentationFormat
                        'Setting up source open presentation template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Numbered List_OpenDocument.odp"
                        'Setting up destination open presentation report 
                        Const strPresentationReport As [String] = "Presentation Reports/Numbered List Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GeneratePieChart(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GeneratePieChartFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GeneratePieChartFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart_DT Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GeneratePieChartFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart_XML Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GeneratePieChartinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GeneratePieChartFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Pie Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Pie Chart_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GeneratePieChartFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Pie Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Pie Chart_DT Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GeneratePieChartFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Pie Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Pie Chart_XML Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GeneratePieChartinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Pie Chart.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Pie Chart Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GeneratePieChartFromDatabaseinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Pie Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Pie Chart_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GeneratePieChartFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Pie Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Pie Chart_DT Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GeneratePieChartFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Pie Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Pie Chart_XML Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GeneratePieChartinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Pie Chart.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Pie Chart Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateScatterChart(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateScatterChartFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateScatterChartFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart_DT Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateScatterChartFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart_DB.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart_XML Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateScatterChartinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart.docx"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateScatterChartFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Scatter Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Scatter Chart_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateScatterChartFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Scatter Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Scatter Chart_DT Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateScatterChartFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Scatter Chart_DB.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Scatter Chart_XML Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateScatterChartinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Scatter Chart.xlsx"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Scatter Chart Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateScatterChartFromDatabaseinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Scatter Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Scatter Chart_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateScatterChartFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Scatter Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Scatter Chart_DT Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersAndOrdersDataDT())
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateScatterChartFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Scatter Chart_DB.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Scatter Chart_XML Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetAllDataFromXML(), "ds")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateScatterChartinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Scatter Chart.pptx"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Scatter Chart Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateSingleRow(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateSingleRowFromDatabaseinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Single Row_DB Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDataDB(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDatabaseinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateSingleRowFromDataSetinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Single Row_DT Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDT(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDataSetinOpenDocumentProcessingFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateSingleRowFromXMLinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Single Row_XML Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerXML(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromXMLinOpenDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateSingleRowinOpenDocumentProcessingFormat
                        'Setting up source open document template
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row_OpenDocument.odt"
                        'Setting up destination open document report 
                        Const strDocumentReport As [String] = "Word Reports/Single Row Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowinOpenDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateSingleRowFromDatabaseinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Single Row_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Single Row_DB Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerDataDB(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDatabaseinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateSingleRowFromDataSetinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Single Row_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Single Row_DT Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerDT(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDataSetinOpenSpreadsheetFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateSingleRowFromXMLinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Single Row_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Single Row_XML Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerXML(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromXMLinOpenSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateSingleRowinOpenSpreadsheetFormat
                        'Setting up source open spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Single Row_OpenDocument.ods"
                        'Setting up destination open document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Single Row Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowinOpenSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateSingleRowFromDatabaseinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Single Row_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Single Row_DB Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerDataDB(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDatabaseinOpenPresentationFormat
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateSingleRowFromDataSetinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Single Row_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Single Row_DT Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerDT(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDataSetinOpenPresentationFormat
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateSingleRowFromXMLinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Single Row_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Single Row_XML Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerXML(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromXMLinOpenPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateSingleRowinOpenPresentationFormat
                        'Setting up source open spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Single Row_OpenDocument.odp"
                        'Setting up destination open document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Single Row Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowinOpenPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub



    End Class
End Namespace

 
