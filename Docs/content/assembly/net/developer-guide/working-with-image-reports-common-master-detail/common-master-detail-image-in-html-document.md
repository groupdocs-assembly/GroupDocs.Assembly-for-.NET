---
id: common-master-detail-image-in-html-document
url: assembly/net/common-master-detail-image-in-html-document
title: Common Master-Detail Image in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail report in HTML Document format.{{< /alert >}}

<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026666554 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026666554 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026666554 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026666554"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#CommonMaster-DetailImageinHTMLDocument-CommonMaster-DetailImageinHTMLDocument">Common Master-Detail Image in HTML Document</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#CommonMaster-DetailImageinHTMLDocument-ReportingRequirement">Reporting Requirement</a></li><li><span class="TOCOutline">1.2</span> <a href="#CommonMaster-DetailImageinHTMLDocument-AddingSyntaxtobeevaluatedbyGroupDocs.AssemblyEngine">Adding Syntax to be evaluated by GroupDocs.Assembly Engine</a></li><li><span class="TOCOutline">1.3</span> <a href="#CommonMaster-DetailImageinHTMLDocument-DownloadCommonMaster-DetailTemplate">Download Common Master-Detail Template</a></li><li><span class="TOCOutline">1.4</span> <a href="#CommonMaster-DetailImageinHTMLDocument-GeneratingTheReport">Generating The Report</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

## Common Master-Detail Image in HTML Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   Report must show customers' picture and name.
*   It must associate the customers with their products.
*   Report must be generated in the HTML Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<<foreach \[in customers\]>>

![](data:image/jpeg;base64,<<[Photo]>>)

**<<\[CustomerName\]>>**

**Products:**

**

<<foreach \[in Order\]>><<\[IndexOf() != 0 ? ", " : ""\]>>

<<\[Product.ProductName\]>>

**

**<</foreach>>**

<</foreach>>

 {{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Common Master-Detail Template

Please download the sample Common master-detail document we created in this article:

*   [Common Master-Detail.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/Common%20Master-Detail.html?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source html template</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>HTML Templates/Common Master-Detail.html<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination html report</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>HTML Reports/Common Master-Detail Report.html<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC7" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Instantiate DocumentAssembler class</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Call AssembleDocument to generate Common Master-Detail Report in html format</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>), <span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>), <span class="pl-smi">DataLayer</span>.<span class="pl-en">PopulateData</span>(), <span class="pl-s"><span class="pl-pds">"</span>customers<span class="pl-pds">"</span></span>);</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC12" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC14" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs-LC16" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/rida-fatima-aspose/31319a8a70569c70ad671106eeb1404b/raw/ed2fa4e17896e90fe19811625eb251e4817c6848/Examples-CSharp-GroupDocs.AssemblyExamples-GenerateReport-GenerateCommonMasterDetailinHtmlFormat.cs) [Examples-CSharp-GroupDocs.AssemblyExamples-GenerateReport-GenerateCommonMasterDetailinHtmlFormat.cs](https://gist.github.com/rida-fatima-aspose/31319a8a70569c70ad671106eeb1404b#file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatecommonmasterdetailinhtmlformat-cs) hosted with ❤ by [GitHub](https://github.com)
