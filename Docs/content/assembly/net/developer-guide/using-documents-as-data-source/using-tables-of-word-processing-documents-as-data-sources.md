---
id: using-tables-of-word-processing-documents-as-data-sources
url: assembly/net/using-tables-of-word-processing-documents-as-data-sources
title: Using Tables of Word Processing Documents as Data Sources
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026666984 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026666984 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026666984 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026666984"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#UsingTablesofWordProcessingDocumentsasDataSources-UsingTablesofWordProcessingDocumentsasDataSources">Using Tables of Word Processing Documents as Data Sources</a></li><li><span class="TOCOutline">2</span> <a href="#UsingTablesofWordProcessingDocumentsasDataSources-TheRecipe">The Recipe</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1</span> <a href="#UsingTablesofWordProcessingDocumentsasDataSources-Download">Download</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1.1</span> <a href="#UsingTablesofWordProcessingDocumentsasDataSources-DataSourceDocument">Data Source Document</a></li><li><span class="TOCOutline">2.1.2</span> <a href="#UsingTablesofWordProcessingDocumentsasDataSources-Template">Template</a></li></ul></li></ul></li><li><span class="TOCOutline">3</span> <a href="#UsingTablesofWordProcessingDocumentsasDataSources-TheCode">The Code</a></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 16.12.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Using Tables of Word Processing Documents as Data Sources

Following classes are added in **[GroupDocs.Assembly.Data](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/)** namespace:

*   [DocumentTable](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttable)
*   [DocumentTableColumn](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablecolumn)
*   [DocumentTableColumnCollection](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablecolumncollection)
*   [DocumentTableOptions](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttableoptions)

## The Recipe

*   Define template and output report documents
*   Assemble a document using the external document table as a data source

### Download

#### Data Source Document

*   [Managers Data.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Word%20DataSource/Managers%20Data.docx?raw=true)

#### Template

*   [Importing Word Processing Table into Presentation.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Importing%20Word%20Processing%20Table%20into%20Presentation.pptx?raw=true)

## The Code

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-importingwordprocessingtableintopresentation-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-k">string</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Presentation Templates/Importing Word Processing Table into Presentation.pptx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">string</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Presentation Reports/Importing Word Processing Table into Presentation_Output.pptx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> Assemble a document using the external document table as a data source.</span></td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC7" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>),</td></tr><tr><td id="file-importingwordprocessingtableintopresentation-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-importingwordprocessingtableintopresentation-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">DataLayer</span>.<span class="pl-en">ImportingWordDocToPresentation</span>(), <span class="pl-s"><span class="pl-pds">"</span>table<span class="pl-pds">"</span></span>));</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/0d70a84f2c42353f38ee844bc6fc0ffb/raw/4b21f113504b2ba0d4ec3865354ab0e59a1db07c/ImportingWordProcessingTableIntoPresentation.cs) [ImportingWordProcessingTableIntoPresentation.cs](https://gist.github.com/GroupDocsGists/0d70a84f2c42353f38ee844bc6fc0ffb#file-importingwordprocessingtableintopresentation-cs) hosted with ❤ by [GitHub](https://github.com)
