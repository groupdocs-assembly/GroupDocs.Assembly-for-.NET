
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports GroupDocs.AssemblyExamples.BusinessLayer
Imports GroupDocs.AssemblyExamples.BusinessLayer.GroupDocs.AssemblyExamples.BusinessLayer
Imports GroupDocs.Assembly

Namespace GroupDocs.AssemblyExamples
    Module Module1

        Sub Main()
            '*
            '             *  Applying product license
            '             *  Please uncomment the statement if you do have license.
            '             

            'CommonUtilities.ApplyLicense()


            '#Region "Generating Bubble Chart Report"
            'Generate a bubble chart report in document processing format
            'GenerateReport.GenerateBubbleChart("document", true)

            'Generate a Bulleted List report in spreadsheet format
            'GenerateReport.GenerateBubbleChart("spreadsheet", true)

            'Generate a Bulleted List report in presentation format
            'GenerateReport.GenerateBubbleChart("presentation", true)
            '#End Region

            '#Region "Generating Bulleted List Report"
            'Generate a Bulleted List report in document processing format
            'GenerateReport.GenerateBulletedList("document", true)

            'Generate a Bulleted List report in spreadsheet format
            'GenerateReport.GenerateBulletedList("spreadsheet", true)

            'Generate a Bulleted List report in presentation format
            'GenerateReport.GenerateBulletedList("presentation", true)
            '#End Region

            '#Region "Generating Chart report with Filtering, Grouping, and Ordering"
            'Generate a Chart report with Filtering, Grouping, and Ordering in document processing format
            'GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("document", true)

            'Generate a Chart report with Filtering, Grouping, and Ordering in spreadsheet format
            'GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("spreadsheet", true)

            'Generate a Chart report with Filtering, Grouping, and Ordering in presentation format
            'GenerateReport.GenerateChartWithFilteringGroupingAndOrdering("presentation", true)
            '#End Region

            '#Region "Generating Common List Report"
            'Generate a Common List Report in document processing format
            'GenerateReport.GenerateCommonList("document", true)

            'Generate a Common List Report in spreadsheet format
            'GenerateReport.GenerateCommonList("spreadsheet", true)

            'Generate a Common List Report in presentation format
            'GenerateReport.GenerateCommonList("presentation", true)
            '#End Region

            '#Region "Generating Common Master-Detail Report"
            'Generate a Common Master-Detail Report in document processing format
            'GenerateReport.GenerateCommonMasterDetail("document", true)

            'Generate a Common Master-Detail Report in spreadsheet format
            'GenerateReport.GenerateCommonMasterDetail("spreadsheet", true)

            'Generate a Common Master-Detail Report in presentation format
            'GenerateReport.GenerateCommonMasterDetail("presentation", true)
            '#End Region

            '#Region "Generating In-Paragraph List Report"
            'Generate a In-Paragraph List Report in document processing format
            'GenerateReport.GenerateInParagraphList("document", true)

            'Generate a In-Paragraph List Report in spreadsheet format
            'GenerateReport.GenerateInParagraphList("spreadsheet", true)

            'Generate a In-Paragraph List Report in presentation format
            'GenerateReport.GenerateInParagraphList("presentation", true)
            '#End Region

            '#Region "Generating In-Table with Alternate Content Report"
            'Generate a In-Table List with Alternate Content Report in document processing format
            'GenerateReport.GenerateInTableListWithAlternateContent("document", true)

            'Generate a In-Table List with Alternate Content Report in spreadsheet format
            'GenerateReport.GenerateInTableListWithAlternateContent("spreadsheet", true)

            'Generate a In-Table List with Alternate Content Report in presentation format
            'GenerateReport.GenerateInTableListWithAlternateContent("presentation", true)
            '#End Region

            '#Region "Generating In-Table with Filtering, Grouping and Ordering Report"
            'Generate a In-Table List with Filtering, Grouping, and Ordering Report in document processing format
            'GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("document", true)

            'Generate a In-Table List with Filtering, Grouping, and Ordering Report in spreadsheet format
            'GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("spreadsheet", true)

            'Generate a In-Table List with Filtering, Grouping, and Ordering Report in presentation format
            'GenerateReport.GenerateInTableListWithFilteringGroupingAndOrdering("presentation", true)
            '#End Region

            '#Region "Generating In-Table List with Highlighted Rows Report"
            'Generate a In-Table List with Highlighted Rows Report in document processing format
            'GenerateReport.GenerateInTableListWithHighlightedRows("document", true)

            'Generate a In-Table List with Highlighted Rows Report in spreadsheet format
            'GenerateReport.GenerateInTableListWithHighlightedRows("spreadsheet", true)

            'Generate a In-Table List with Highlighted Rows Report in presentation format
            'GenerateReport.GenerateInTableListWithHighlightedRows("presentation", true)
            '#End Region

            '#Region "Generating In-Table List Report"
            'Generate a In-Table List Report in document processing format
            'GenerateReport.GenerateInTableList("document", true)

            'Generate a In-Table List Report in spreadsheet format
            'GenerateReport.GenerateInTableList("spreadsheet", true)

            'Generate a In-Table List Report in presentation format
            'GenerateReport.GenerateInTableList("presentation", true)
            '#End Region

            '#Region "Generating In-Table Master-Detail Report"
            'Generate a In-Table Master-Detail Report in document processing format
            'GenerateReport.GenerateInTableMasterDetail("document", true)

            'Generate a In-Table Master-Detail Report in spreadsheet format
            'GenerateReport.GenerateInTableMasterDetail("spreadsheet", false)

            'Generate a In-Table Master-Detail Report in presentation format
            'GenerateReport.GenerateInTableMasterDetail("presentation", true)
            '#End Region

            '#Region "Generating Multicolored Number List Report"
            'Generate a Multicolored Numbered List Report in document processing format
            'GenerateReport.GenerateMulticoloredNumberedList("document", true)

            'Generate a Multicolored Numbered List Report in spreadsheet format
            'GenerateReport.GenerateMulticoloredNumberedList("spreadsheet", true)

            'Generate a Multicolored Numbered List Report in presentation format
            'GenerateReport.GenerateMulticoloredNumberedList("presentation", true)
            '#End Region

            '#Region "Generating Numbered List Report"
            'Generate a Numbered List Report in document processing format
            'GenerateReport.GenerateNumberedList("document", true)

            'Generate a Numbered List Report in spreadsheet format
            'GenerateReport.GenerateNumberedList("spreadsheet", true)

            'Generate a Numbered List Report in presentation format
            'GenerateReport.GenerateNumberedList("presentation", true)
            '#End Region

            '#Region "Generating Pie Chart Report"
            'Generate a Pie Chart Report in document processing format
            'GenerateReport.GeneratePieChart("document", True)
            '
            'Generate a Pie Chart Report in spreadsheet format
            'GenerateReport.GeneratePieChart("spreadsheet", true)

            'Generate a Pie Chart Report in presentation format
            'GenerateReport.GeneratePieChart("presentation", true)
            '#End Region

            '#Region "Generating Scatter Chart Report"
            'Generate a Scatter Chart Report in document processing format
            'GenerateReport.GenerateScatterChart("document", true)

            'Generate a Scatter Chart Report in spreadsheet format
            'GenerateReport.GenerateScatterChart("spreadsheet", true)

            'Generate a Scatter Chart Report in presentation format
            GenerateReport.GenerateScatterChart("presentation", True)
            '#End Region

            '#Region "Generating Single Row Report"
            'Generate a Single Row Report in document processing format
            GenerateReport.GenerateSingleRow("document", True)

            'Generate a Single Row Report in spreadsheet format
            GenerateReport.GenerateSingleRow("spreadsheet", True)

            'Generate a Single Row Report in presentation format
            GenerateReport.GenerateSingleRow("presentation", True)
            '#End Region


        End Sub
    End Module
End Namespace

