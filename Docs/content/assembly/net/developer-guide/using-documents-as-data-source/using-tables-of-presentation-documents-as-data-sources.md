---
id: using-tables-of-presentation-documents-as-data-sources
url: assembly/net/using-tables-of-presentation-documents-as-data-sources
title: Using Tables of Presentation Documents as Data Sources
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026666977 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026666977 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026666977 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026666977"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#UsingTablesofPresentationDocumentsasDataSources-UsingTablesofPresentationDocumentsasDataSources">Using Tables of Presentation Documents as Data Sources</a></li><li><span class="TOCOutline">2</span> <a href="#UsingTablesofPresentationDocumentsasDataSources-TheRecipe">The Recipe</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1</span> <a href="#UsingTablesofPresentationDocumentsasDataSources-Download">Download</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1.1</span> <a href="#UsingTablesofPresentationDocumentsasDataSources-DataSourceDocument">Data Source Document</a></li><li><span class="TOCOutline">2.1.2</span> <a href="#UsingTablesofPresentationDocumentsasDataSources-Template">Template</a></li></ul></li></ul></li><li><span class="TOCOutline">3</span> <a href="#UsingTablesofPresentationDocumentsasDataSources-TheCode">The Code</a></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 16.12.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Using Tables of Presentation Documents as Data Sources

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

*   [Managers Data.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Presentation%20DataSource/Managers%20Data.pptx?raw=true)

#### Template

*   [Using Presentation as Table of Data.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Using%20Presentation%20as%20Table%20of%20Data.pptx?raw=true)

## The Code

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-usepresentationtableasdatasource-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-usepresentationtableasdatasource-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-usepresentationtableasdatasource-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-k">string</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Presentation Templates/Using Presentation as Table of Data.pptx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-usepresentationtableasdatasource-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">string</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Presentation Reports/Using Presentation as Table of Data_Output.pptx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-usepresentationtableasdatasource-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> Assemble a document using the external document table as a data source.</span></td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-usepresentationtableasdatasource-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-usepresentationtableasdatasource-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-usepresentationtableasdatasource-cs-LC7" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>),</td></tr><tr><td id="file-usepresentationtableasdatasource-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-usepresentationtableasdatasource-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">DataLayer</span>.<span class="pl-en">PresentationData</span>(), <span class="pl-s"><span class="pl-pds">"</span>table<span class="pl-pds">"</span></span>));</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/b69cdd70c2f91362b2c6fdddd7bea8a5/raw/0cb688aafccd6187e6436610903ff2b1fc8bed36/UsePresentationTableAsDataSource.cs) [UsePresentationTableAsDataSource.cs](https://gist.github.com/GroupDocsGists/b69cdd70c2f91362b2c6fdddd7bea8a5#file-usepresentationtableasdatasource-cs) hosted with ❤ by [GitHub](https://github.com)
