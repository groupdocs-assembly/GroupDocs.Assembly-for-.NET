
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports GroupDocs.AssemblyExamples.BusinessLayer
Imports GroupDocs.AssemblyExamples.BusinessLayer.GroupDocs.AssemblyExamples.BusinessLayer
Imports GroupDocs.Assembly
Imports System.Reflection

Namespace GroupDocs.AssemblyExamples
    Module Module1

        Sub Main()
            'ExStart:ApplyingLicense
            '*
            '             *  Applying product license
            '             *  Please uncomment the statement if you do have license.
            '             

            'CommonUtilities.ApplyLicense()
            'ExEnd:ApplyingLicense

            '#Region "Generating Bubble Chart Report"
            'Generate a bubble chart report in document processing format
            GenerateReport.GenerateBubbleChart("document", False, True)

            'Generate a Bulleted List report in spreadsheet format
            GenerateReport.GenerateBubbleChart("spreadsheet", False, True)

            'Generate a Bulleted List report in presentation format
            GenerateReport.GenerateBubbleChart("presentation", False, True)
            '#End Region

            '#Region "Generating Bulleted List Report"
            'Generate a Bulleted List report in document processing format
            GenerateReport.GenerateBulletedList("document", False, True)

            'Generate a Bulleted List report in spreadsheet format
            GenerateReport.GenerateBulletedList("spreadsheet", False, True)

            'Generate a Bulleted List report in presentation format
            GenerateReport.GenerateBulletedList("presentation", False, True)
            '#End Region

            '#Region "Generating Chart report with Filtering, Grouping, and Ordering"
            'Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document", False, True)

            'Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet", False, True)

            'Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation", False, True)
            '#End Region

            '#Region "Generating Common List Report"
            'Generate a Common List Report in document processing format
            GenerateReport.GenerateCommonList("document", False, True)

            'Generate a Common List Report in spreadsheet format
            GenerateReport.GenerateCommonList("spreadsheet", False, True)

            'Generate a Common List Report in presentation format
            GenerateReport.GenerateCommonList("presentation", False, True)
            '#End Region

            '#Region "Generating Common Master-Detail Report"
            'Generate a Common Master-Detail Report in document processing format
            GenerateReport.GenerateCommonMasterDetail("document", False, True)

            'Generate a Common Master-Detail Report in spreadsheet format
            GenerateReport.GenerateCommonMasterDetail("spreadsheet", False, True)

            'Generate a Common Master-Detail Report in presentation format
            GenerateReport.GenerateCommonMasterDetail("presentation", False, True)
            '#End Region

            '#Region "Generating In-Paragraph List Report"
            'Generate a In-Paragraph List Report in document processing format
            GenerateReport.GenerateInParagraphList("document", False, True)

            'Generate a In-Paragraph List Report in spreadsheet format
            GenerateReport.GenerateInParagraphList("spreadsheet", False, True)

            'Generate a In-Paragraph List Report in presentation format
            GenerateReport.GenerateInParagraphList("presentation", False, True)
            '#End Region

            '#Region "Generating In-Table with Alternate Content Report"
            'Generate a In-Table List with Alternate Content Report in document processing format
            GenerateReport.GenerateInTableListWithAlternateContent("document", False, True)

            'Generate a In-Table List with Alternate Content Report in spreadsheet format
            GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet", False, True)

            'Generate a In-Table List with Alternate Content Report in presentation format
            GenerateReport.GenerateInTableListWithAlternateContent("presentation", False, True)
            '#End Region

            '#Region "Generating In-Table with Filtering, Grouping and Ordering Report"
            'Generate a In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document", False, True)

            'Generate a In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet", False, True)

            'Generate a In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation", False, True)
            '#End Region

            '#Region "Generating In-Table List with Highlighted Rows Report"
            'Generate a In-Table List with Highlighted Rows Report in document processing format
            GenerateReport.GenerateInTableListWithHighlightedRows("document", False, True)

            'Generate a In-Table List with Highlighted Rows Report in spreadsheet format
            GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet", False, True)

            'Generate a In-Table List with Highlighted Rows Report in presentation format
            GenerateReport.GenerateInTableListWithHighlightedRows("presentation", False, True)
            '#End Region

            '#Region "Generating In-Table List Report"
            'Generate a In-Table List Report in document processing format
            GenerateReport.GenerateInTableList("document", False, True)

            'Generate a In-Table List Report in spreadsheet format
            GenerateReport.GenerateInTableList("spreadsheet", False, True)

            'Generate a In-Table List Report in presentation format
            GenerateReport.GenerateInTableList("presentation", False, True)
            '#End Region

            '#Region "Generating In-Table Master-Detail Report"
            'Generate a In-Table Master-Detail Report in document processing format
            GenerateReport.GenerateInTableMasterDetail("document", False, True)

            'Generate a In-Table Master-Detail Report in spreadsheet format
            GenerateReport.GenerateInTableMasterDetail("spreadsheet", False, True)

            'Generate a In-Table Master-Detail Report in presentation format
            GenerateReport.GenerateInTableMasterDetail("presentation", False, True)
            '#End Region

            '#Region "Generating Multicolored Number List Report"
            'Generate a Multicolored Numbered List Report in document processing format
            GenerateReport.GenerateMulticoloredNumberedList("document", False, True)

            'Generate a Multicolored Numbered List Report in spreadsheet format
            GenerateReport.GenerateMulticoloredNumberedList("spreadsheet", False, True)

            'Generate a Multicolored Numbered List Report in presentation format
            GenerateReport.GenerateMulticoloredNumberedList("presentation", False, True)
            '#End Region

            '#Region "Generating Numbered List Report"
            'Generate a Numbered List Report in document processing format
            GenerateReport.GenerateNumberedList("document", False, True)

            'Generate a Numbered List Report in spreadsheet format
            GenerateReport.GenerateNumberedList("spreadsheet", False, True)

            'Generate a Numbered List Report in presentation format
            GenerateReport.GenerateNumberedList("presentation", False, True)
            '#End Region

            '#Region "Generating Pie Chart Report"
            'Generate a Pie Chart Report in document processing format
            GenerateReport.GeneratePieChart("document", False, True)
            '
            'Generate a Pie Chart Report in spreadsheet format
            GenerateReport.GeneratePieChart("spreadsheet", False, True)

            'Generate a Pie Chart Report in presentation format
            GenerateReport.GeneratePieChart("presentation", False, True)
            '#End Region

            '#Region "Generating Scatter Chart Report"
            'Generate a Scatter Chart Report in document processing format
            GenerateReport.GenerateScatterChart("document", False, True)

            'Generate a Scatter Chart Report in spreadsheet format
            GenerateReport.GenerateScatterChart("spreadsheet", False, True)

            'Generate a Scatter Chart Report in presentation format
            GenerateReport.GenerateScatterChart("presentation", False, True)
            '#End Region

            '#Region "Generating Single Row Report"
            'Generate a Single Row Report in document processing format
            GenerateReport.GenerateSingleRow("document", False, True)

            'Generate a Single Row Report in spreadsheet format
            'GenerateReport.GenerateSingleRow("spreadsheet", False, True)

            'Generate a Single Row Report in presentation format
            'GenerateReport.GenerateSingleRow("presentation", False, True)
            '#End Region

        End Sub

    End Module
End Namespace

