---
id: migrating-to-groupdocs-assembly-3-0-0-or-higher
url: assembly/net/migrating-to-groupdocs-assembly-3-0-0-or-higher
title: Migrating to GroupDocs.Assembly 3.0.0 or Higher
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026664009 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026664009 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026664009 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026664009"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#MigratingtoGroupDocs.Assembly3.0.0orHigher-ReasonstoMigrate">Reasons to Migrate</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#MigratingtoGroupDocs.Assembly3.0.0orHigher-Namespace(s)">Namespace(s)</a></li><li><span class="TOCOutline">1.2</span> <a href="#MigratingtoGroupDocs.Assembly3.0.0orHigher-SupportofMultipleDataSources">Support of Multiple Data Sources</a></li><li><span class="TOCOutline">1.3</span> <a href="#MigratingtoGroupDocs.Assembly3.0.0orHigher-EasyTemplateSyntax">Easy Template Syntax</a></li><li><span class="TOCOutline">1.4</span> <a href="#MigratingtoGroupDocs.Assembly3.0.0orHigher-TemplateEditing">Template Editing</a></li><li><span class="TOCOutline">1.5</span> <a href="#MigratingtoGroupDocs.Assembly3.0.0orHigher-ComfyCode">Comfy Code</a></li></ul></li></ul></div></div></div></td><td valign="top">&nbsp;</td></tr></tbody></table>

## Reasons to Migrate

Just to endeavor the report generation task more smoothly, we released **GroupDocs.Assembly for .NET 3.0.0**.

### Namespace(s)

In **GroupDocs.Assembly for .NET 3.0.0** only a single **using GroupDocs.Assembly;** namespace is required to generate reports in any of the supported formats. Whereas, in **GroupDocs.Assembly for .NET 1.3.0** to generate a report in any of the supported format a separate namespace of that format is required to be added, given are the namespaces:

1.  using Groupdocs.Assembly;
2.  using Groupdocs.Assembly.Cells;
3.  using Groupdocs.Assembly.Data;
4.  using Groupdocs.Assembly.Exceptions;
5.  using Groupdocs.Assembly.Pdf;
6.  using Groupdocs.Assembly.Slides;
7.  using Groupdocs.Assembly.Words;

### Support of Multiple Data Sources

**GroupDocs.Assembly for .NET 3.0.0** supports following data sources:

*   Database
*   XML
*   JSON
*   OData
*   Custom .NET Objects

### Easy Template Syntax

**Next Generation GroupDocs.Assembly for .NET** engine supports the underlying C# syntax for LINQ queries which is actually shorter than LINQ. This basically means that the developers can use the familiar and well documented C# syntax to write data binding/traversal queries right in the document templates. As a result, developers can enjoy many benefits including short and concise reporting syntax and binding to any type of supported data source including business objects.

### Template Editing

In **GroupDocs.Assembly for .NET 1.3.0** user need to edit the template through code. However, template editing in **GroupDocs.Assembly for .NET 3.0.0** is quite straightforward. Font size, text alignment and other similar context is dealt on the template end.

### Comfy Code

Code to generate reports has been shrunk to fewer lines in **GroupDocs.Assembly for .NET 3.0.0**. A user does not need to write the bulk of lines of code to generate big reports. Instead, he/she will have to write 2 to 3 lines of code to generate business documents and reports.

{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer{{< /alert >}}

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-k">static</span> <span class="pl-k">void</span> <span class="pl-en">Main</span>(<span class="pl-k">string</span>[] <span class="pl-smi">args</span>)</td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC3" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source word template</span></td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-en">FileStream</span> <span class="pl-smi">template</span> <span class="pl-k">=</span> <span class="pl-smi">File</span>.<span class="pl-en">OpenRead</span>(<span class="pl-s"><span class="pl-pds">"</span>../../../../Data/Samples/Source/Pie Chart.docx<span class="pl-pds">"</span></span>);</td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination word report</span></td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC7" class="blob-code blob-code-inner js-file-line"><span class="pl-en">FileStream</span> <span class="pl-smi">output</span> <span class="pl-k">=</span> <span class="pl-smi">File</span>.<span class="pl-en">Create</span>(<span class="pl-s"><span class="pl-pds">"</span>../../../../Data/Samples/Destination/Pie Chart Report.docx<span class="pl-pds">"</span></span>);</td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC8" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Generate the report</span></td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">doc</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Beside data source and output path it takes dataSourceObject and dataSourceString</span></td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">doc</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">template</span>, <span class="pl-smi">output</span>, <span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>( <span class="pl-en">GetProductsData</span>(), <span class="pl-s"><span class="pl-pds">"</span>customers<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-generating-report-in-groupdocs-assembly-3-0-0-cs-LC13" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/73103f412435effbc30b8f2c5b2366ba/raw/56e89d86612b108cdb3a311c6e6982f0b71b2e72/generating%20report%20in%20groupdocs.assembly%203.0.0.cs) [generating report in groupdocs.assembly 3.0.0.cs](https://gist.github.com/GroupDocsGists/73103f412435effbc30b8f2c5b2366ba#file-generating-report-in-groupdocs-assembly-3-0-0-cs) hosted with ❤ by [GitHub](https://github.com)
