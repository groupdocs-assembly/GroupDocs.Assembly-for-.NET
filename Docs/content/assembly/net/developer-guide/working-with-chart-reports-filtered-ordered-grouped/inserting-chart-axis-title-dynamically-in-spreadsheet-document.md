---
id: inserting-chart-axis-title-dynamically-in-spreadsheet-document
url: assembly/net/inserting-chart-axis-title-dynamically-in-spreadsheet-document
title: Inserting Chart Axis Title Dynamically in Spreadsheet Document
weight: 7
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026664846 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026664846 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026664846 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026664846"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-ColumnChartinMicrosoftExcelDocument">Column Chart in Microsoft Excel Document</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-CreatingaColumnChart">Creating a Column Chart</a></li><li><span class="TOCOutline">1.2</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-ReportingRequirement">Reporting Requirement</a></li><li><span class="TOCOutline">1.3</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-AddingSyntaxtobeEvaluatedbyGroupDocs.AssemblyEngine">Adding Syntax to be Evaluated by GroupDocs.Assembly Engine</a><ul class="toc-indentation"><li><span class="TOCOutline">1.3.1</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-ChartTitle">Chart Title</a></li><li><span class="TOCOutline">1.3.2</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-ChartData">Chart Data</a></li></ul></li><li><span class="TOCOutline">1.4</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-DownloadTemplate">Download Template</a></li><li><span class="TOCOutline">1.5</span> <a href="#InsertingChartAxisTitleDynamicallyinSpreadsheetDocument-GeneratingTheReport">Generating The Report</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 18.1 or greater.{{< /alert >}}

## Column Chart in Microsoft Excel Document

### Creating a Column Chart

Following steps, you can create a column chart in MS Excel 2013:

1.  Create a new Workbook
2.  Click the "Insert" tab, and then click "Insert Column Chart" icon in the Charts group to view the drop-down list
3.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your Worksheet
4.  Edit the Worksheet with your data to update the chart
5.  Save your Document

### Reporting Requirement

As a report developer, you are required to share orders quantity of the customers dynamically with the following key requirements:

*   A report must show the quantity of sales/orders
*   Sales/orders quantity must be represented by Quarters
*   It must associate order quantity with the corresponding customer
*   A report must be generated in the Spreadsheet Document

### Adding Syntax to be Evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
<<[title]>><<foreach [in orders
.Where(c => c.OrderDate.Year == 2015)
.GroupBy(c => c.Customer)
.OrderBy(g => g.Key.CustomerName)]>><<x [Key.CustomerName]>>

```

#### Chart Data

**Legend Entries**

```csharp
="1st Quarter<<y [Where(c => c.OrderDate.Month >= 1 && c.OrderDate.Month <= 3).Sum(c => c.ProductQuantity)]>>"
="2nd Quarter<<y [Where(c => c.OrderDate.Month >= 4 && c.OrderDate.Month <= 6).Sum(c => c.ProductQuantity)]>>"
="3rd Quarter<<y [Where(c => c.OrderDate.Month >= 7 && c.OrderDate.Month <= 9).Sum(c => c.ProductQuantity)]>>"
="4th Quarter<<y [Where(c => c.OrderDate.Month >= 10 && c.OrderDate.Month <= 12).Sum(c => c.ProductQuantity)]>>"

```

### Download Template

Please download the sample Chart with Filtering, Grouping, and Ordering with Dynamic Title document we created in this article:

*   [Chart Template.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_Dynamic_Title.xlsx) (Template for CustomObject and JSON examples) 

### Generating The Report

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source open document template</span></td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Spreadsheet Templates/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.xlsx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination open document report</span></td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Spreadsheet Reports/Chart with Filtering, Grouping, and Ordering_Dynamic_Title.xlsx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC7" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Instantiate DocumentAssembler class</span></td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-k">string</span> <span class="pl-smi">title</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Total Order Quantity by Quarters<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format</span></td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>),</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC14" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>( <span class="pl-smi">DataLayer</span>.<span class="pl-en">GetOrdersData</span>(),<span class="pl-s"><span class="pl-pds">"</span>orders<span class="pl-pds">"</span></span>),</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">title</span>,<span class="pl-s"><span class="pl-pds">"</span>title<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC16" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC18" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L19" class="blob-num js-line-number" data-line-number="19"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC19" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-dynamicchartaxistitlespreadsheet-cs-L20" class="blob-num js-line-number" data-line-number="20"></td><td id="file-dynamicchartaxistitlespreadsheet-cs-LC20" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/251fa6d077fd45b1222a6f352c4329b2/raw/ccd6ffadcb37bcad7410b59040137145e262ff9d/DynamicChartAxisTitleSpreadSheet.cs) [DynamicChartAxisTitleSpreadSheet.cs](https://gist.github.com/GroupDocsGists/251fa6d077fd45b1222a6f352c4329b2#file-dynamicchartaxistitlespreadsheet-cs) hosted with ❤ by [GitHub](https://github.com)
