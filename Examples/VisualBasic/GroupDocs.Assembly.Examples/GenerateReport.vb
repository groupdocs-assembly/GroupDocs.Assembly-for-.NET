
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

        Public Shared Sub GenerateBubbleChart(strDocumentFormat As String, isDatabase As Boolean)

            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateBubbleChartFromDatabaseinDocumentProcessingFormat
                        'setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'initialize object of DocumentAssembler class 
                            'Call AssembleDocument to generate Bubble Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateBubbleChartinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bubble Chart.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Bubble Chart Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartinDocumentProcessingFormat
                        End Try
                    End If

                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateBubbleChartFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bubble Chart_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bubble Chart_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateBubbleChartinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bubble Chart.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bubble Chart Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateBubbleChartFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bubble Chart_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bubble Chart_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateBubbleChartinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bubble Chart.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bubble Chart Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bubble Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBubbleChartinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateBulletedList(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateBulletedListFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateBulletedListinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Bulleted List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Bulleted List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateBulletedListFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bulleted List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bulleted List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateBulletedListinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Bulleted List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Bulleted List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateBulletedListFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bulleted List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bulleted List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateBulletedListinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Bulleted List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Bulleted List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Bulleted List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateBulletedListinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateChartWithFilteringGroupingAndOrdering(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderinginDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Chart with Filtering, Grouping, and Ordering.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Chart with Filtering, Grouping, and Ordering Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderinginSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderingFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateChartWithFilteringGroupingAndOrderinginPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Chart with Filtering, Grouping, and Ordering.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Chart with Filtering, Grouping, and Ordering Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateCommonList(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateCommonListFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Common List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonListinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Common List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateCommonListFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonListinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateCommonListFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonListinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonListinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateCommonMasterDetail(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateCommonMasterDetailFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common Master-Detail Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonMasterDetailinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Common Master-Detail.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Common Master-Detail Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common Master-Detail Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateCommonMasterDetailFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common Master-Detail_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common Master-Detail_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonMasterDetailinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Common Master-Detail.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Common Master-Detail Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateCommonMasterDetailFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common Master-Detail_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common Master-Detail_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateCommonMasterDetailinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Common Master-Detail.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Common Master-Detail Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Common List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateCommonMasterDetailinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateInParagraphList(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInParagraphListFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateInParagraphListinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Paragraph List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Paragraph List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInParagraphListFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Paragraph List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Paragraph List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInParagraphListinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Paragraph List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Paragraph List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInParagraphListFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Paragraph List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Paragraph List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInParagraphListinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Paragraph List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Paragraph List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Paragraph List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInParagraphListinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableListWithAlternateContent(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithAlternateContentinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Alternate Content.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Alternate Content Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithAlternateContentinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Alternate Content.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Alternate Content Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithAlternateContentFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Alternate Content_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithAlternateContentinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Alternate Content.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Alternate Content Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Alternate Content Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithAlternateContentinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableListWithFilteringGroupingAndOrdering(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Filtering, Grouping, and Ordering.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Filtering, Grouping, and Ordering Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Filtering, Grouping, and Ordering.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Filtering, Grouping, and Ordering Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderingFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithFilteringGroupingAndOrderinginPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Filtering, Grouping, and Ordering.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Filtering, Grouping, and Ordering Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Filtering, Grouping, and Ordering Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateInTableListWithHighlightedRows(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List with Highlighted Rows.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List with Highlighted Rows Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in document format
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
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinSpreadsheetDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List with Highlighted Rows.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List with Highlighted Rows Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListWithHighlightedRowsFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListWithHighlightedRowsinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List with Highlighted Rows.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List with Highlighted Rows Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List with Highlighted Rows Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListWithHighlightedRowsinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateInTableList(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If True Then
                        'ExStart:GenerateInTableListFromDatabaseinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDatabaseinDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateInTableListinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListinDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableListFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableListFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableListinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableListinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
        Public Shared Sub GenerateInTableMasterDetail(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateInTableMasterDetailFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableMasterDetailinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/In-Table Master-Detail.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/In-Table Master-Detail Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateInTableMasterDetailFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table Master-Detail_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableMasterDetailinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/In-Table Master-Detail.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/In-Table Master-Detail Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateInTableMasterDetailFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table Master-Detail_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table Master-Detail_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateInTableMasterDetailinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/In-Table Master-Detail.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/In-Table Master-Detail Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate In-Table Master-Detail Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateInTableMasterDetailinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateMulticoloredNumberedList(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateMulticoloredNumberedListFromDatabaseinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDatabaseinDocumentProcessingDocument
                        End Try
                    Else
                        'ExStart:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Multicolored Numbered List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Multicolored Numbered List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateMulticoloredNumberedListFromDatabaseinSpreadsheetDocument
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Multicolored Numbered List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDatabaseinSpreadsheetDocument
                        End Try
                    Else
                        'ExStart:GenerateMulticoloredNumberedListinSpreadsheetDocument
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Multicolored Numbered List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Multicolored Numbered List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListinSpreadsheetDocument
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateMulticoloredNumberedListFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Multicolored Numbered List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Multicolored Numbered List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateMulticoloredNumberedListinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Multicolored Numbered List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Multicolored Numbered List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Multicolored Numbered List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateMulticoloredNumberedListinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateNumberedList(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateNumberedListFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateNumberedListinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Numbered List.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Numbered List Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateNumberedListFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Numbered List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Numbered List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateNumberedListinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Numbered List.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Numbered List Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateNumberedListFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Numbered List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Numbered List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsDataDB(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateNumberedListinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Numbered List.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Numbered List Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Numbered List Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetProductsData(), "products")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateNumberedListinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GeneratePieChart(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GeneratePieChartFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GeneratePieChartinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Pie Chart.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Pie Chart Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GeneratePieChartFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Pie Chart_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Pie Chart_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GeneratePieChartinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Pie Chart.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Pie Chart Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GeneratePieChartFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Pie Chart_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Pie Chart_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomersDataDB(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GeneratePieChartinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Pie Chart.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Pie Chart Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Pie Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.PopulateData(), "customers")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GeneratePieChartinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateScatterChart(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateScatterChartFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart_DB.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart_DB Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateScatterChartinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Scatter Chart.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Scatter Chart Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateScatterChartFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Scatter Chart_DB.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Scatter Chart_DB Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateScatterChartinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Scatter Chart.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Scatter Chart Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateScatterChartFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Scatter Chart_DB.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Scatter Chart_DB Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersDataDB(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateScatterChartinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Scatter Chart.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Scatter Chart Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Scatter Chart Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetOrdersData(), "orders")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateScatterChartinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub

        Public Shared Sub GenerateSingleRow(strDocumentFormat As String, isDatabase As Boolean)
            Select Case strDocumentFormat
                Case "document"
                    If isDatabase Then
                        'ExStart:GenerateSingleRowFromDatabaseinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Single Row Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetSingleCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDatabaseinDocumentProcessingFormat
                        End Try
                    Else
                        'ExStart:GenerateSingleRowinDocumentProcessingFormat
                        'Setting up source document template
                        Const strDocumentTemplate As [String] = "Word Templates/Single Row.docx"
                        'Setting up destination document report 
                        Const strDocumentReport As [String] = "Word Reports/Single Row Report.docx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in document format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate), CommonUtilities.SetDestinationDocument(strDocumentReport), DataLayer.GetCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowinDocumentProcessingFormat
                        End Try
                    End If
                    Exit Select

                Case "spreadsheet"
                    If isDatabase Then
                        'ExStart:GenerateSingleRowFromDatabaseinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Single Row.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Single Row Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetSingleCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDatabaseinSpreadsheetFormat
                        End Try
                    Else
                        'ExStart:GenerateSingleRowinSpreadsheetFormat
                        'Setting up source spreadsheet template
                        Const strSpreadsheetTemplate As [String] = "Spreadsheet Templates/Single Row.xlsx"
                        'Setting up destination document report 
                        Const strSpreadsheetReport As [String] = "Spreadsheet Reports/Single Row Report.xlsx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in spreadsheet format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strSpreadsheetTemplate), CommonUtilities.SetDestinationDocument(strSpreadsheetReport), DataLayer.GetCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowinSpreadsheetFormat
                        End Try
                    End If
                    Exit Select

                Case "presentation"
                    If isDatabase Then
                        'ExStart:GenerateSingleRowFromDatabaseinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Single Row.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Single Row Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetSingleCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowFromDatabaseinPresentationFormat
                        End Try
                    Else
                        'ExStart:GenerateSingleRowinPresentationFormat
                        'Setting up source spreadsheet template
                        Const strPresentationTemplate As [String] = "Presentation Templates/Single Row.pptx"
                        'Setting up destination document report 
                        Const strPresentationReport As [String] = "Presentation Reports/Single Row Report.pptx"
                        Try
                            'Instantiate DocumentAssembler class
                            Dim assembler As New DocumentAssembler()
                            'Call AssembleDocument to generate Single Row Report in presentation format
                            assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strPresentationTemplate), CommonUtilities.SetDestinationDocument(strPresentationReport), DataLayer.GetCustomerData(), "customer")
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                            'ExEnd:GenerateSingleRowinPresentationFormat
                        End Try
                    End If
                    Exit Select
            End Select
        End Sub
    End Class
End Namespace

 
