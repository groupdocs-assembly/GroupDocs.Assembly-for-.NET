---
id: removing-selective-chart-series-dynamically
url: assembly/net/removing-selective-chart-series-dynamically
title: Removing Selective Chart Series Dynamically
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026664802 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026664802 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026664802 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026664802"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#RemovingSelectiveChartSeriesDynamically-ColumnChartinMicrosoftWordDocument">Column Chart in Microsoft Word Document</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#RemovingSelectiveChartSeriesDynamically-CreatingaColumnChart">Creating a Column Chart</a></li><li><span class="TOCOutline">1.2</span> <a href="#RemovingSelectiveChartSeriesDynamically-ReportingRequirement">Reporting Requirement</a></li><li><span class="TOCOutline">1.3</span> <a href="#RemovingSelectiveChartSeriesDynamically-AddingSyntaxtobeevaluatedbyGroupDocs.AssemblyEngine">Adding Syntax to be evaluated by GroupDocs.Assembly Engine</a><ul class="toc-indentation"><li><span class="TOCOutline">1.3.1</span> <a href="#RemovingSelectiveChartSeriesDynamically-ChartTitle">Chart Title</a></li><li><span class="TOCOutline">1.3.2</span> <a href="#RemovingSelectiveChartSeriesDynamically-ChartData(Excel)">Chart Data (Excel)</a></li></ul></li><li><span class="TOCOutline">1.4</span> <a href="#RemovingSelectiveChartSeriesDynamically-DownloadTemplate">Download Template</a></li><li><span class="TOCOutline">1.5</span> <a href="#RemovingSelectiveChartSeriesDynamically-GeneratingTheReport">Generating The Report</a><ul class="toc-indentation"><li><span class="TOCOutline">1.5.1</span> <a href="#RemovingSelectiveChartSeriesDynamically-CustomObjects">Custom Objects</a></li></ul></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.11 or greater.{{< /alert >}}

## Column Chart in Microsoft Word Document

### Creating a Column Chart

Practising the following steps, you can insert a Column Chart in MS Word 2013:

1.  Click in the document where you want to insert the chart, click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box.
2.  Select "Column" in the sidebar, you will see a gallery of charts.
3.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your document.
4.  Edit the Worksheet with your data to update the chart. See [Chart Data (Excel)]({{< ref "assembly/net/developer-guide/working-with-chart-reports-filtered-ordered-grouped/removing-selective-chart-series-dynamically.md" >}}).
5.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share orders quantity of the customers dynamically with the following key requirements:

*   A report must show the quantity of sales/orders.
*   Sales/orders quantity must represented by Quarters.
*   It must associate order quantity with the corresponding customer.
*   The report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Quantity by Quarters<<foreach [in orders
.Where(c => c.OrderDate.Year == 2015)
.GroupBy(c => c.Customer)
.OrderBy(g => g.Key.CustomerName)]>><<x [Key.CustomerName]>>

```

#### Chart Data (Excel)

|   | 1st Quarter<<y[Where(c => c.OrderDate.Month>= 1 && c.OrderDate.Month <= 3).Sum(c => c.ProductQuantity)]>><<removeif [mode== 1]>> | 2nd Quarter<<y[Where(c => c.OrderDate.Month>= 4 && c.OrderDate.Month <= 6).Sum(c => c.ProductQuantity)]>><<removeif [mode==2]>> | 3rd Quarter<<y[Where(c => c.OrderDate.Month>= 7 && c.OrderDate.Month <= 9).Sum(c => c.ProductQuantity)]>> | 4th Quarter<<y[Where(c => c.OrderDate.Month>= 10 && c.OrderDate.Month <= 12).Sum(c => c.ProductQuantity)]>> |
| --- | --- | --- | --- | --- |
| Category 1 | 4.3 | 2.4 | 2 | 3 |
| Category 2 | 2.5 | 4.4 | 2 | 2 |
| Category 3 | 3.5 | 1.8 | 3 | 5 |
| Category 4 | 4.5 | 2.8 | 5 | 2 |

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Template

Please download the sample Chart with Filtering, Grouping, and Ordering document we created in this article:

*   [Chart Template.docx](https://github.com/aliahmedgroupdocs/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_RemoveIf.docx) (Template for CustomObject and JSON examples) 

### Generating The Report

#### Custom Objects

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-removeselectivechartseries-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-removeselectivechartseries-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-removeselectivechartseries-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-removeselectivechartseries-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source open document template</span></td></tr><tr><td id="file-removeselectivechartseries-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-removeselectivechartseries-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Templates/Chart with Filtering, Grouping, and Ordering_RemoveIf.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-removeselectivechartseries-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-removeselectivechartseries-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination open document report</span></td></tr><tr><td id="file-removeselectivechartseries-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-removeselectivechartseries-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/Chart with Filtering, Grouping, and Ordering_RemoveIf.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-removeselectivechartseries-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-removeselectivechartseries-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-removeselectivechartseries-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-removeselectivechartseries-cs-LC7" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-removeselectivechartseries-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-removeselectivechartseries-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Instantiate DocumentAssembler class</span></td></tr><tr><td id="file-removeselectivechartseries-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-removeselectivechartseries-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-removeselectivechartseries-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-removeselectivechartseries-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-k">int</span> <span class="pl-smi">mode</span> <span class="pl-k">=</span> <span class="pl-c1">2</span>;</td></tr><tr><td id="file-removeselectivechartseries-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-removeselectivechartseries-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Call AssembleDocument to generate Chart report with Filtering, Grouping, and Ordering in document format</span></td></tr><tr><td id="file-removeselectivechartseries-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-removeselectivechartseries-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-removeselectivechartseries-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-removeselectivechartseries-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>),</td></tr><tr><td id="file-removeselectivechartseries-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-removeselectivechartseries-cs-LC14" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">DataLayer</span>.<span class="pl-en">GetOrdersData</span>(),<span class="pl-s"><span class="pl-pds">"</span>orders<span class="pl-pds">"</span></span>),</td></tr><tr><td id="file-removeselectivechartseries-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-removeselectivechartseries-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">mode</span>,<span class="pl-s"><span class="pl-pds">"</span>mode<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-removeselectivechartseries-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-removeselectivechartseries-cs-LC16" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-removeselectivechartseries-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-removeselectivechartseries-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-removeselectivechartseries-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-removeselectivechartseries-cs-LC18" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-removeselectivechartseries-cs-L19" class="blob-num js-line-number" data-line-number="19"></td><td id="file-removeselectivechartseries-cs-LC19" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-removeselectivechartseries-cs-L20" class="blob-num js-line-number" data-line-number="20"></td><td id="file-removeselectivechartseries-cs-LC20" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/b44bbcc5a2ca4b9f6fd0c854e792902f/raw/06c067543310832fbd424a101e074ad7096303c3/RemoveSelectiveChartSeries.cs) [RemoveSelectiveChartSeries.cs](https://gist.github.com/GroupDocsGists/b44bbcc5a2ca4b9f6fd0c854e792902f#file-removeselectivechartseries-cs) hosted with ❤ by [GitHub](https://github.com)
