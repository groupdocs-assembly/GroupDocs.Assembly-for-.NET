---
id: working-with-csv-data-sources
url: assembly/net/working-with-csv-data-sources
title: Working with CSV Data Sources
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026667242 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026667242 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026667242 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026667242"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#WorkingwithCSVDataSources-TreatingsimpleCSVdata"><span style="color: rgb(0, 0, 0);">Treating simple CSV data</span></a></li><li><span class="TOCOutline">2</span> <a href="#WorkingwithCSVDataSources-Configuredatasourcetoreadcolumnnames">Configure data source to read column names</a></li><li><span class="TOCOutline">3</span> <a href="#WorkingwithCSVDataSources-Download">&nbsp;Download</a><ul class="toc-indentation"><li><span class="TOCOutline">3.1</span> <a href="#WorkingwithCSVDataSources-DataSourceDocument">Data Source Document</a></li><li><span class="TOCOutline">3.2</span> <a href="#WorkingwithCSVDataSources-Template">Template</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 19.10 or later releases.{{< /alert >}}

To access CSV data while building a report, you can pass a [CsvDataSource](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/csvdatasource)instance to the assembler as a data source. Using of *CsvDataSource* enables you to work with typed values rather than just strings in template documents. Although CSV as a format does not define a way to store values of types other than strings, *CsvDataSource* is capable to recognize values of the following types by their string representations:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?

{{< alert style="warning" >}}For recognition of data types to work, string representations of corresponding values must be formed using invariant culture settings.{{< /alert >}}

### Treating simple CSV data

In template documents, a [CsvDataSource ](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/csvdatasource)instance should be treated in the same way as if it was a *DataTable* instance (see "[Working with *DataTable* and *DataView* Objects](https://docs.groupdocs.com/display/assemblynet/Template+Syntax+-+Part+1+of+2#TemplateSyntax-Part1of2-DataTableandDataViewObjects)" for more information) as shown in the following example.

Suppose we have CSV data like:

John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm

By using the template like following:

  

<<foreach \[in persons\]>>Name: <<\[Column1\]>>, Age: <<\[Column2\]>>, Date
of Birth: <<\[Column3\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Column2)\]>>

  

The results will be produced like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format and the extension method involving arithmetic in the template document becomes possible, because text values of Column3 and Column2 are automatically converted to DateTime? and Int32? respectively.{{< /alert >}}

### Configure data source to read column names

By default, [CsvDataSource ](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/csvdatasource)uses column names such as "*Column1*", "*Column2*", and so on, as you can see from the previous example. However, *CsvDataSource* can be configured to read column names from the first line of CSV data as shown in the following example.

Suppose we have CSV data like:

Name,Age,Birth
John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm

By using the template like following:

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of
Birth: <<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

The code should be written like:

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-simplecsvds_demo_19-10-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-simplecsvds_demo_19-10-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-simplecsvds_demo_19-10-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source document template (Email or Word Document)</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-simplecsvds_demo_19-10-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Text Templates/CsvDatasetDemo.txt<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-simplecsvds_demo_19-10-cs-LC4" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-simplecsvds_demo_19-10-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination for reports</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-simplecsvds_demo_19-10-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/SimpleCsvDSDemo Out.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-simplecsvds_demo_19-10-cs-LC7" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-simplecsvds_demo_19-10-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination Markdown reports</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-simplecsvds_demo_19-10-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDataSource</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Excel DataSource/Person.csv<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-simplecsvds_demo_19-10-cs-LC10" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-simplecsvds_demo_19-10-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up CSV data load options</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-simplecsvds_demo_19-10-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-en">CsvDataLoadOptions</span> <span class="pl-smi">options</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">CsvDataLoadOptions</span>(<span class="pl-c1">true</span>);</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-simplecsvds_demo_19-10-cs-LC13" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-simplecsvds_demo_19-10-cs-LC14" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>initialize data source</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-simplecsvds_demo_19-10-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-en">CsvDataSource</span> <span class="pl-smi">dataSource</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">CsvDataSource</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetDataSourceDocument</span>(<span class="pl-smi">strDataSource</span>), <span class="pl-smi">options</span>);</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-simplecsvds_demo_19-10-cs-LC16" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-simplecsvds_demo_19-10-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Assemble the document</span></td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-simplecsvds_demo_19-10-cs-LC18" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L19" class="blob-num js-line-number" data-line-number="19"></td><td id="file-simplecsvds_demo_19-10-cs-LC19" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L20" class="blob-num js-line-number" data-line-number="20"></td><td id="file-simplecsvds_demo_19-10-cs-LC20" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>),</td></tr><tr><td id="file-simplecsvds_demo_19-10-cs-L21" class="blob-num js-line-number" data-line-number="21"></td><td id="file-simplecsvds_demo_19-10-cs-LC21" class="blob-code blob-code-inner js-file-line"><span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">dataSource</span>, <span class="pl-s"><span class="pl-pds">"</span>persons<span class="pl-pds">"</span></span>));</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/f59b71ca7e6345f9ac817297f3707f5f/raw/b8d24431d8e7a39534ef353b5a353718874e0372/SimpleCsvDS_Demo_19.10.cs) [SimpleCsvDS\_Demo\_19.10.cs](https://gist.github.com/GroupDocsGists/f59b71ca7e6345f9ac817297f3707f5f#file-simplecsvds_demo_19-10-cs) hosted with ❤ by [GitHub](https://github.com)

The results will be produced like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

Also, you can use *CsvDataLoadOptions* to customize the following characters playing special roles while loading CSV data:

*   Value separator (the default is comma)
*   Single-line comment start (the default is sharp)
*   Quotation mark enabling to use other special characters within a value (the default is double quotes)

###  Download

#### Data Source Document

*   [Person.csv](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/XML%20DataSource/Managers.xml?raw=true)

#### Template

*   [Demo Template.txt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)
