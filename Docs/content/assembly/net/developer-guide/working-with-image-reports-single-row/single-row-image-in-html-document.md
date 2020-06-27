---
id: single-row-image-in-html-document
url: assembly/net/single-row-image-in-html-document
title: Single Row Image in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Single Row Image report in HTML Document format.{{< /alert >}}

<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026665504 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026665504 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026665504 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026665504"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#SingleRowImageinHTMLDocument-SingleRowImageinHTMLDocument">Single Row Image in HTML Document</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#SingleRowImageinHTMLDocument-ReportingRequirement">Reporting Requirement</a></li><li><span class="TOCOutline">1.2</span> <a href="#SingleRowImageinHTMLDocument-AddingSyntaxtobeevaluatedbyGroupDocs.AssemblyEngine">Adding Syntax to be evaluated by GroupDocs.Assembly Engine</a></li><li><span class="TOCOutline">1.3</span> <a href="#SingleRowImageinHTMLDocument-DownloadSingleRowImageTemplate">Download Single Row Image Template</a></li><li><span class="TOCOutline">1.4</span> <a href="#SingleRowImageinHTMLDocument-GeneratingTheReport">Generating The Report</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

## Single Row Image in HTML Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent information of first single customer with the following key requirements:

*   Report must show image of the customer
*   It must show Name and Contact Number of the customer
*   Report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
<table>
<tr>
<td>
<img height="100" width="100" src="data:image/jpeg;base64,&lt;&lt;[customer.Photo]&gt;&gt;"/>
</td>
<td>
<table>
<tr>
<td>
<b>Name:</b>
</td>
<td>
<b>&lt;&lt;[customer.CustomerName]&gt;&gt;</b>  
</td>
</tr>
<tr>
<td>
<b>Contact Number:</b>
</td>
<td>
&lt;&lt;[customer.CustomerContactNumber]&gt;&gt;
</td>
</tr>
</table>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Single Row Image Template

Please download the single row image document we created in this article:

*   [Single Row.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Numbered%20List.txt?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer{{< /alert >}}

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source html template</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strHtmlTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>HTML Templates/Single Row.html<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination html report</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strHtmlReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>HTML Reports/Single Row Report.html<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC7" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Instantiate DocumentAssembler class</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Call AssembleDocument to generate Single Row Report in html format</span></td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strHtmlTemplate</span>),</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strHtmlReport</span>),</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">DataLayer</span>.<span class="pl-en">GetCustomerData</span>(), <span class="pl-s"><span class="pl-pds">"</span>customer<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC14" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC16" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs-LC18" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/9bbee318cdfc682815c7284ada1d94bf/raw/9a03b54263cfc4d70514c99197e49fa1a1c667af/Examples-CSharp-GroupDocs.AssemblyExamples-GenerateReport-GenerateSingleRowinHtmlFormat.cs) [Examples-CSharp-GroupDocs.AssemblyExamples-GenerateReport-GenerateSingleRowinHtmlFormat.cs](https://gist.github.com/GroupDocsGists/9bbee318cdfc682815c7284ada1d94bf#file-examples-csharp-groupdocs-assemblyexamples-generatereport-generatesinglerowinhtmlformat-cs) hosted with ❤ by [GitHub](https://github.com)
