
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

        Public Shared Sub GenerateBubbleChart(strDocumentFormat As String)

            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateBubbleChartinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateBubbleChartinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateBubbleChartinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateBulletedList(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateBulletedListinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateBulletedListinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateBulletedListinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateChartWithFilteringGroupingAndOrdering(strDocumentFormat As String)

            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateChartWithFilteringGroupingAndOrderinginPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateCommonList(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateCommonListinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateCommonListinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateCommonListinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateCommonMasterDetail(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateCommonMasterDetailinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateCommonMasterDetailinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateCommonMasterDetailinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateInParagraphList(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateInParagraphListinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateInParagraphListinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateInParagraphListinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateInTableListWithAlternateContent(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateInTableListWithAlternateContentinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateInTableListWithAlternateContentinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateInTableListWithAlternateContentinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateInTableListWithFilteringGroupingAndOrdering(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginDocumentProcessingDocument
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateInTableListWithFilteringGroupingAndOrderinginPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateInTableListWithHighlightedRows(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateInTableListWithHighlightedRowsinDocumentProcessingDocument
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateInTableListWithHighlightedRowsinSpreadsheetDocument
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateInTableListWithHighlightedRowsinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateInTableList(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateInTableListinDocumentProcessingDocument
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateInTableListinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateInTableListinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateInTableMasterDetail(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateInTableMasterDetailinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateInTableMasterDetailinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateInTableMasterDetailinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateMulticoloredNumberedList(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateMulticoloredNumberedListinDocumentProcessingDocument
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateMulticoloredNumberedListinSpreadsheetDocument
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateMulticoloredNumberedListinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateNumberedList(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateNumberedListinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateNumberedListinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateNumberedListinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GeneratePieChart(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GeneratePieChartinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GeneratePieChartinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GeneratePieChartinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateScatterChart(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateScatterChartinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateScatterChartinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateScatterChartinPresentationFormat
                    Exit Select
            End Select
        End Sub


        Public Shared Sub GenerateSingleRow(strDocumentFormat As String)
            Select Case strDocumentFormat
                Case "document"
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
                    End Try
                    'ExEnd:GenerateSingleRowinDocumentProcessingFormat
                    Exit Select

                Case "spreadsheet"
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
                    End Try
                    'ExEnd:GenerateSingleRowinSpreadsheetFormat
                    Exit Select

                Case "presentation"
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
                    End Try
                    'ExEnd:GenerateSingleRowinPresentationFormat
                    Exit Select
            End Select
        End Sub
    End Class
End Namespace


