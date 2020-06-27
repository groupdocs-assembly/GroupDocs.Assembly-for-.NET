---
id: working-with-xml-data-sources
url: assembly/net/working-with-xml-data-sources
title: Working with XML Data Sources
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026667264 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026667264 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026667264 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026667264"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#WorkingwithXMLDataSources-Treatingthetop-levelXMLelementcontainsonlyasequenceofelementsofthesametype">Treating the top-level XML element contains only a sequence of elements of the same type</a></li><li><span class="TOCOutline">2</span> <a href="#WorkingwithXMLDataSources-Treatingthetop-levelXMLelementcontainsattributesornestedelementsofdifferenttypes">Treating the top-level XML element contains attributes or nested elements of different types</a></li><li><span class="TOCOutline">3</span> <a href="#WorkingwithXMLDataSources-Thecompleteexample">The complete example</a><ul class="toc-indentation"><li><span class="TOCOutline">3.1</span> <a href="#WorkingwithXMLDataSources-XML">XML</a></li><li><span class="TOCOutline">3.2</span> <a href="#WorkingwithXMLDataSources-Templatedocument"><span>Template document</span></a></li><li><span class="TOCOutline">3.3</span> <a href="#WorkingwithXMLDataSources-Sourcecode"><span>Source code</span></a></li><li><span class="TOCOutline">3.4</span> <a href="#WorkingwithXMLDataSources-Resultdocument">Result document</a></li></ul></li><li><span class="TOCOutline">4</span> <a href="#WorkingwithXMLDataSources-Download">Download</a><ul class="toc-indentation"><li><span class="TOCOutline">4.1</span> <a href="#WorkingwithXMLDataSources-DataSourceDocument">Data Source Document</a></li><li><span class="TOCOutline">4.2</span> <a href="#WorkingwithXMLDataSources-Template">Template</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 19.10 or later releases.{{< /alert >}}

To access XML data while building a report, you can use facilities of *DataSet* to read XML into it and then pass it to the assembler as a data source. However, if your scenario does not permit to specify XML schema while loading XML into *DataSet*, all attributes and text values of XML elements are loaded as strings then. Thus, it becomes impossible, for example, to use arithmetic operations on numbers or to specify custom date-time and numeric formats to output orresponding values, because all of them are treated as strings. To overcome this limitation, you can pass an *XmlDataSource* instance to the assembler as a data source instead. Even when XML schema is not provided, *XmlDataSource* is capable to recognize values of the following types by their string representations:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?

{{< alert style="warning" >}}For recognition of data types to work, string representations of corresponding attributes and text values of XML elements must be formed using invariant culture settings.{{< /alert >}}

### Treating the top-level XML element contains only a sequence of elements of the same type

In template documents, if a top-level XML element contains only a sequence of elements of the same type, an *XmlDataSource* instance should be treated in the same way as if it was a *DataTable *instance (see "[Working with *DataTable* and *DataView* Objects](https://docs.groupdocs.com/display/assemblynet/Template+Syntax+-+Part+1+of+2#TemplateSyntax-Part1of2-DataTableandDataViewObjects)" for more information).

Suppose we have XML data like:

<Persons>
   <Person>
      <Name>John Doe</Name>
      <Age>30</Age>
      <Birth>1989-04-01 4:00:00 pm</Birth>
   </Person>
   <Person>
      <Name>Jane Doe</Name>
      <Age>27</Age>
      <Birth>1992-01-31 07:00:00 am</Birth>
   </Person>
   <Person>
      <Name>John Smith</Name>
      <Age>51</Age>
      <Birth>1968-03-08 1:00:00 pm</Birth>
   </Person>
</Persons>

By using the template like following:

  

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth: 
<<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

  

The results will be produced like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format and the extension method involving arithmetic in the template document becomes possible, because text values of Birth and Age XML elements are automatically converted to DateTime? and Int32? respectively even in the absence of XML schema.{{< /alert >}}

### Treating the top-level XML element contains attributes or nested elements of different types

If a top-level XML element contains attributes or nested elements of different types, an *XmlDataSource* instance should be treated in template documents in the same way as if it was a *DataRow* instance (see "[Working with *DataRow* and *DataRowView* Objects](https://docs.groupdocs.com/display/assemblynet/Template+Syntax+-+Part+1+of+2#TemplateSyntax-Part1of2-DataTableandDataViewObjects)" for more information) as shown in the following example.

Suppose we have XML data like:

<Person>
   <Name>John Doe</Name>
   <Age>30</Age>
   <Birth>1989-04-01 4:00:00 pm</Birth>
   <Child>Ann Doe</Child>
   <Child>Charles Doe</Child>
</Person>

By using the template like following:

Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth:
<<\[Birth\]:"dd.MM.yyyy">>
Children:
<<foreach \[in Child\]>><<\[Child\_Text\]>>
<</foreach>>

The results will be produced like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Children:
Ann Doe
Charles Doe

{{< alert style="warning" >}}To reference a sequence of repeated simple-type XML elements with the same name, the elements’ name itself (for example, "Child") should be used in a template document, whereas the same name with the "_Text" suffix (for example, "Child_Text") should be used to reference the text value of one of these elements.{{< /alert >}}

### The complete example

The following example sums up typical scenarios involving nested complex-type XML elements.

#### XML

<Managers>
   <Manager>
      <Name>John Smith</Name>
      <Contract>
         <Client>
            <Name>A Company</Name>
         </Client>
         <Price>1200000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>B Ltd.</Name>
         </Client>
         <Price>750000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>C &amp; D</Name>
         </Client>
         <Price>350000</Price>
      </Contract>
   </Manager>
   <Manager>
      <Name>Tony Anderson</Name>
      <Contract>
         <Client>
            <Name>E Corp.</Name>
         </Client>
         <Price>650000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>F &amp; Partners</Name>
         </Client>
         <Price>550000</Price>
      </Contract>
   </Manager>
   <Manager>
      <Name>July James</Name>
      <Contract>
         <Client>
            <Name>G &amp; Co.</Name>
         </Client>
         <Price>350000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>H Group</Name>
         </Client>
         <Price>250000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>I &amp; Sons</Name>
            43
         </Client>
         <Price>100000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>J Ent.</Name>
         </Client>
         <Price>100000</Price>
      </Contract>
   </Manager>
</Managers>

#### Template document

<<foreach \[in managers\]>>Manager: <<\[Name\]>>
Contracts:
<<foreach \[in Contract\]>>- <<\[Client.Name\]>> ($<<\[Price\]>>)
<</foreach>>
<</foreach>>

#### Source code

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-simplexmlds_demo_19-10-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-simplexmlds_demo_19-10-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-simplexmlds_demo_19-10-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source document template (Email or Word Document)</span></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-simplexmlds_demo_19-10-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Templates/SimpleDatasetDemo.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-simplexmlds_demo_19-10-cs-LC4" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-simplexmlds_demo_19-10-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination for reports</span></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-simplexmlds_demo_19-10-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/SimpleXMLDSDemo Out.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-simplexmlds_demo_19-10-cs-LC7" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-simplexmlds_demo_19-10-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination Markdown reports</span></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-simplexmlds_demo_19-10-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDataSource</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>XML DataSource/Managers.xml<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-simplexmlds_demo_19-10-cs-LC10" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-simplexmlds_demo_19-10-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>initialize data source</span></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-simplexmlds_demo_19-10-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-en">XmlDataSource</span> <span class="pl-smi">dataSource</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">XmlDataSource</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetDataSourceDocument</span>(<span class="pl-smi">strDataSource</span>));</td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-simplexmlds_demo_19-10-cs-LC13" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-simplexmlds_demo_19-10-cs-LC14" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> Assemble the document</span></td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-simplexmlds_demo_19-10-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-simplexmlds_demo_19-10-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-simplexmlds_demo_19-10-cs-LC16" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>), <span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>), <span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">dataSource</span>, <span class="pl-s"><span class="pl-pds">"</span>managers<span class="pl-pds">"</span></span>));</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/ddac17d9a69c71495b2931ca6c7e8ce7/raw/2134d8e13665abf5a0cb1b28b3c6eae57998f0e1/SimpleXMLDS_Demo_19.10.cs) [SimpleXMLDS\_Demo\_19.10.cs](https://gist.github.com/GroupDocsGists/ddac17d9a69c71495b2931ca6c7e8ce7#file-simplexmlds_demo_19-10-cs) hosted with ❤ by [GitHub](https://github.com)

  

#### Result document

Manager: John Smith
Contracts:
- A Company ($1200000)
- B Ltd. ($750000)
- C & D ($350000)
Manager: Tony Anderson
Contracts:
- E Corp. ($650000)
- F & Partners ($550000)
Manager: July James
Contracts:
- G & Co. ($350000)
- H Group ($250000)
- I & Sons ($100000)
- J Ent. ($100000)

### Download

#### Data Source Document

*   [Managers.xml](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/XML%20DataSource/Managers.xml?raw=true)

#### Template

*   [Demo Template.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)
