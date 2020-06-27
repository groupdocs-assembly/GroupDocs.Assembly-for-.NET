---
id: individual-series-point-coloring-in-word-processing-document
url: assembly/net/individual-series-point-coloring-in-word-processing-document
title: Individual Series Point Coloring in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026666802 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026666802 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026666802 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026666802"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-IndividualSeriesPointColoringinWordProcessingDocument">Individual Series Point Coloring in Word Processing Document</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-ReportingRequirement">Reporting Requirement</a></li><li><span class="TOCOutline">1.2</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-AddingSyntaxtobeevaluatedbyGroupDocs.AssemblyEngine">Adding Syntax to be evaluated by GroupDocs.Assembly Engine</a><ul class="toc-indentation"><li><span class="TOCOutline">1.2.1</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-ChartTitle">Chart Title</a></li><li><span class="TOCOutline">1.2.2</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-ChartData(Excel)">Chart Data (Excel)</a></li></ul></li><li><span class="TOCOutline">1.3</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-DownloadPieChartTemplate">Download Pie Chart Template</a></li><li><span class="TOCOutline">1.4</span> <a href="#IndividualSeriesPointColoringinWordProcessingDocument-GeneratingTheReport">Generating The Report</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}This feature is supported by version 18.5 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Individual Series Point Coloring in Word Processing Document

Please follow below steps to can create Pie Chart in MS Word 2013.

1.  Click in the document where you want to insert the chart, click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box.
2.  Select "Pie".
3.  Preview "Pie" and press OK to insert the chart and Worksheet template to your document.
4.  Edit the Worksheet with your data to update the chart. See [Chart Data (Excel)]({{< ref "assembly/net/developer-guide/working-with-individual-series-point-coloring/individual-series-point-coloring-in-word-processing-document.md" >}})
5.  Save the template.

### Reporting Requirement

As a report developer, you are required to share your customers' orders details dynamically with the following key requirements:

*   The report must show information on a Pie Chart
*   It must indicate customer name with value (price of the orders purchased)
*   The report must be generated in the Word Processing Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Price<<foreach [in customers]>>
<<x [CustomerName]>>

```

#### Chart Data (Excel)

|   | Total Order Price<<y [Order.Sum(c => c.Price)]>><<pointColor [Color] >> |
| --- | --- |
| A | 8.2 |
| B | 3.2 |
| C | 1.5 |
| D | 1.2 |

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Pie Chart Template

Please download the sample Pie Chart document we created in this article:

*   [Dynamic Chart Point Series Color.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Dynamic%20Chart%20Point%20Series%20Color.docx) (Template for CustomObject and JSON examples) 

### Generating The Report

For a chart with dynamic data, you can set colors of individual chart series points dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series with points to be colored dynamically, define corresponding color expressions in names of these series using **pointColor** tags having the following syntax:
    
    ```csharp
    <<pointColor [color_expression]>>
    ```
    

A color expression must return a value of one of the following types:

*   A string containing the name of a known color, that is, the case-insensitive name of a member of the [KnownColor](https://msdn.microsoft.com/en-us/library/system.drawing.knowncolor(v=vs.110).aspx) enumeration such as "red"
*   An integer value defining RGB (red, green, blue) components of the color such as 0xFFFF00 (yellow)
*   A value of the [Color](http://msdn.microsoft.com/en-us/library/system.drawing.color(v=vs.110).aspx) type

Following code snippet generates the report:

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-dynamicchartseriespointcolor-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-dynamicchartseriespointcolor-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-dynamicchartseriespointcolor-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>setting up source</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-dynamicchartseriespointcolor-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Templates/Dynamic Chart Point Series Color.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-dynamicchartseriespointcolor-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-dynamicchartseriespointcolor-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/Dynamic Chart Point Series Color.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-dynamicchartseriespointcolor-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-dynamicchartseriespointcolor-cs-LC7" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-dynamicchartseriespointcolor-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Instantiate DocumentAssembler class</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-dynamicchartseriespointcolor-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();<span class="pl-c"><span class="pl-c">//</span>initialize object of DocumentAssembler class</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-dynamicchartseriespointcolor-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Call AssembleDocument to generate Pie Chart report in document format</span></td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-dynamicchartseriespointcolor-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-dynamicchartseriespointcolor-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>),</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-dynamicchartseriespointcolor-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">DataLayer</span>.<span class="pl-en">GetCustomerDataFromJson</span>(), <span class="pl-s"><span class="pl-pds">"</span>customers<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-dynamicchartseriespointcolor-cs-LC14" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-dynamicchartseriespointcolor-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-dynamicchartseriespointcolor-cs-LC16" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-dynamicchartseriespointcolor-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-dynamicchartseriespointcolor-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-dynamicchartseriespointcolor-cs-LC18" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/fb1400ea59b283001465944f5a2780d8/raw/ee45261d4b989ba294afe0c5c1a3437c51c35ac4/DynamicChartSeriesPointColor.cs) [DynamicChartSeriesPointColor.cs](https://gist.github.com/GroupDocsGists/fb1400ea59b283001465944f5a2780d8#file-dynamicchartseriespointcolor-cs) hosted with ❤ by [GitHub](https://github.com)
