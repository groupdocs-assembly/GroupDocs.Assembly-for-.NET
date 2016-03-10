﻿
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
       
        Public Shared Sub GenerateBubbleChart(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateBubbleChartFromJsoninWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart.docx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart_Json Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromJsoninWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateBubbleChartFromJsoninExcel
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Bubble Chart.xlsx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Bubble Chart_Json Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromJsoninExcel
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
                    ElseIf isJson Then
                        'ExStart:GenerateBubbleChartFromJsoninPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Bubble Chart.pptx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Bubble Chart_Json Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromJsoninPresentation
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
        Public Shared Sub GenerateBulletedList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateBulletedListFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bulleted list report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateBulletedListFromJsoninOpenExcel
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Bulleted List_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Bulleted List_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate bulleted list in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromJsoninOpenExcel
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
                    ElseIf isJson Then
                        'ExStart:GenerateBulletedListFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Bulleted List_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Bulleted List_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate bulleted list in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromJsoninOpenPresentation
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
        Public Shared Sub GenerateChartWithFilteringGroupingAndOrdering(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromJsoninWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering_Json Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromJsoninWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromJsoninSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_Json Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromJsoninSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromJsoninPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_Json Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Chart with Filtering, Grouping, and Ordering report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromJsoninPresentation
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
        Public Shared Sub GenerateCommonList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateCommonListReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Common List_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Common List_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Common List report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateCommonListReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Common List_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Common List_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Common List report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateCommonListReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Common List_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Common List_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Common List report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateCommonMasterDetail(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateCommonMasterDetailReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Common master-detail report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateCommonMasterDetailReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Common Master-Detail_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Common Master-Detail_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Common master-detail report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateCommonMasterDetailReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Common Master-Detail_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Common Master-Detail_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Common master-detail report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateInParagraphList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateInParagraphListReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Paragraph List report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateInParagraphListReportFromJsoninOpenExcel
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/In-Paragraph List_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/In-Paragraph List_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Paragraph List report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListReportFromJsoninOpenExcel
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
                    ElseIf isJson Then
                        'ExStart:GenerateInParagraphListReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/In-Paragraph List_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/In-Paragraph List_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Paragraph List report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateInTableListWithAlternateContent(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithAlternateContentReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Alternate Content report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithAlternateContentReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithAlternateContentReportFromJsoninOpenExcel
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Alternate Content report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithAlternateContentReportFromJsoninOpenExcel
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithAlternateContentReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/In-Table List with Alternate Content_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Alternate Content report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithAlternateContentReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateInTableListWithFilteringGroupingAndOrdering(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDataSetinDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromXMLinDocumentProcessingDocument
                        End Try
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenWord
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithFilteringGroupingAndOrderingReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateInTableListWithHighlightedRows(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromXMLinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinDocumentProcessingDocument
                        End Try
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenWord
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinSpreadsheetDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinSpreadsheetDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDataSetinSpreadsheetDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDataSetinSpreadsheetDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromXMLinSpreadsheetDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromXMLinSpreadsheetDocument
                        End Try
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenSpreadsheet
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
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
                            'ExEnd:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListwithHighlightedRowsReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateInTableList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListFromDatabaseinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListFromDatabaseinDocumentProcessingDocument
                        End Try
                    ElseIf isDataSet Then
                        'ExStart:GenerateInTableListFromDataSetinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListFromDataSetinDocumentProcessingDocument
                        End Try
                    ElseIf isDataSourceXML Then
                        'ExStart:GenerateInTableListFromXMLinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListFromXMLinDocumentProcessingDocument
                        End Try
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListReportFromJsoninOpenWord
                        End Try
                    Else
                        'ExStart:GenerateInTableListinDocumentProcessingDocument
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
                            'ExEnd:GenerateInTableListinDocumentProcessingDocument
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/In-Table List_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/In-Table List_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableListReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/In-Table List_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/In-Table List_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table List report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateInTableMasterDetail(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableMasterDetailReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table Master-Detail report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableMasterDetailReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/In-Table Master-Detail_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table Master-Detail report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateInTableMasterDetailReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/In-Table Master-Detail_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/In-Table Master-Detail_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate In-Table Master-Detail report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateMulticoloredNumberedList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateMulticoloredNumberedListReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Multicolored Numbered List report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateMulticoloredNumberedListReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Multicolored Numbered List_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Multicolored Numbered List report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateMulticoloredNumberedListReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Multicolored Numbered List_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Multicolored Numbered List_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Multicolored Numbered List report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateNumberedList(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateNumberedListReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Numbered List report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateNumberedListReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Numbered List_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Numbered List_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Numbered List report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateNumberedListReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Numbered List_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Numbered List_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Numbered List report in open presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataJson(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListReportFromJsoninOpenPresentation
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
        Public Shared Sub GeneratePieChart(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GeneratePieChartReportFromJsoninWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart.docx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart_Json Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Pie Chart report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartReportFromJsoninWord
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
                    ElseIf isJson Then
                        'ExStart:GeneratePieChartReportFromJsoninSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Pie Chart.xlsx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Pie Chart_Json Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Pie Chart report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartReportFromJsoninSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GeneratePieChartReportFromJsoninPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Pie Chart.pptx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Pie Chart_Json Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Pie Chart report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerDataFromJson(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartReportFromJsoninPresentation
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
        Public Shared Sub GenerateScatterChart(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateScatterChartReportFromJsoninWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart.docx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart_Json Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Scatter Chart report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartReportFromJsoninWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateScatterChartReportFromJsoninSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Scatter Chart.xlsx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Scatter Chart_Json Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Scatter Chart report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartReportFromJsoninSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateScatterChartReportFromJsoninPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Scatter Chart.pptx"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Scatter Chart_Json Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Scatter Chart report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerOrderDataFromJson(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartReportFromJsoninPresentation
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
        Public Shared Sub GenerateSingleRow(strDocumentFormat As String, isDatabase As Boolean, isDataSet As Boolean, isDataSourceXML As Boolean, isJson As Boolean)
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
                    ElseIf isJson Then
                        'ExStart:GenerateSingleRowReportFromJsoninOpenWord
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row_OpenDocument.odt"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Word Reports/Single Row_OpenDocument_Json Report.odt"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Single Row report in open document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDataJson(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowReportFromJsoninOpenWord
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
                    ElseIf isJson Then
                        'ExStart:GenerateSingleRowReportFromJsoninOpenSpreadsheet
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Spreadsheet Templates/Single Row_OpenDocument.ods"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Spreadsheet Reports/Single Row_OpenDocument_Json Report.ods"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Single Row report in open spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDataJson(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowReportFromJsoninOpenSpreadsheet
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
                    ElseIf isJson Then
                        'ExStart:GenerateSingleRowReportFromJsoninOpenPresentation
                        'setting up source 
                        Const strDocumentTemplate As [String] = "Presentation Templates/Single Row_OpenDocument.odp"
                        'Setting up destination 
                        Const strDocumentReport As [String] = "Presentation Reports/Single Row_OpenDocument_Json Report.odp"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Single Row report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerDataJson(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowReportFromJsoninOpenPresentation
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
        Public Shared Sub GenerateReportLazilyAndRecursively()
            'ExStart:GeneratingReportbyRecursivelyandLazilyAccessingtheData
            'Setting up source open document template
            Const strDocumentTemplate As [String] = "Word Templates/Lazy And Recursive.docx"
            'Setting up destination open document report 
            Const strDocumentReport As [String] = "Word Reports/Lazy And Recursive Report.docx"
            Try
                'Instantiate DynamicEntity class
                Dim dentity As New DynamicEntity(Guid.NewGuid())
                'Instantiate DocumentAssembler class
                Dim assembler As New DocumentAssembler()
                'Call AssembleDocument to generate Single Row Report in open document format
                assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), dentity, "root")
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
            'ExEnd:GeneratingReportbyRecursivelyandLazilyAccessingtheData
        End Sub

    End Class
End Namespace

 
