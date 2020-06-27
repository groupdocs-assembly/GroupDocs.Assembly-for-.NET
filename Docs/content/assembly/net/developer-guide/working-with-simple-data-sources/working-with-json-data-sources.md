---
id: working-with-json-data-sources
url: assembly/net/working-with-json-data-sources
title: Working with JSON Data Sources
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026667253 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026667253 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026667253 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026667253"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#WorkingwithJSONDataSources-Treatingthetoplevelarraysorobjectshavingarray">Treating the top level arrays or objects having array</a></li><li><span class="TOCOutline">2</span> <a href="#WorkingwithJSONDataSources-Treatingtheobjectsattoplevel">Treating the objects at top level</a></li><li><span class="TOCOutline">3</span> <a href="#WorkingwithJSONDataSources-Thecompleteexample">The complete example</a><ul class="toc-indentation"><li><span class="TOCOutline">3.1</span> <a href="#WorkingwithJSONDataSources-JSON">JSON</a></li><li><span class="TOCOutline">3.2</span> <a href="#WorkingwithJSONDataSources-Templatedocument"><span>Template document</span></a></li><li><span class="TOCOutline">3.3</span> <a href="#WorkingwithJSONDataSources-Sourcecode"><span>Source code</span></a></li><li><span class="TOCOutline">3.4</span> <a href="#WorkingwithJSONDataSources-Resultdocument">Result document</a></li></ul></li><li><span class="TOCOutline">4</span> <a href="#WorkingwithJSONDataSources-Download">Download</a><ul class="toc-indentation"><li><span class="TOCOutline">4.1</span> <a href="#WorkingwithJSONDataSources-DataSourceDocument">Data Source Document</a></li><li><span class="TOCOutline">4.2</span> <a href="#WorkingwithJSONDataSources-Template">Template</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td></tr></tbody></table>

{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 19.10 or later releases.{{< /alert >}}

To access JSON data while building a report, the GroupDocs.Assembly API introduces *JsonDataSource* class. You can pass its instance to the assembler as a data source. Using of *JsonDataSource* enables you to work with typed values of JSON elements in template

documents. For more convenience, the set of simple JSON types is extended as follows:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?
*   String

### Treating the top level arrays or objects having array

In template documents, if a top-level JSON element is an array or an object having only one property of an array type, a JsonDataSource instance should be treated in the same way as if it was a DataTable instance (see "[Working with DataTable and DataView Objects](https://docs.groupdocs.com/display/assemblynet/Template+Syntax+-+Part+1+of+2#TemplateSyntax-Part1of2-DataTableandDataViewObjects)" for more information) .

Suppose we have Json data like:

\[ 
   { 
      Name:"John Doe",
      Age:30,
      Birth:"1989-04-01 4:00:00 pm"
   },
   { 
      Name:"Jane Doe",
      Age:27,
      Birth:"1992-01-31 07:00:00 am"
   },
   { 
      Name:"John Smith",
      Age:51,
      Birth:"1968-03-08 1:00:00 pm"
   }
\]

or alternative JSON like:

{ 
   Persons:\[ 
      { 
         Name:"John Doe",
         Age:30,
         Birth:"1989-04-01 4:00:00 pm"
      },
      { 
         Name:"Jane Doe",
         Age:27,
         Birth:"1992-01-31 07:00:00 am"
      },
      { 
         Name:"John Smith",
         Age:51,
         Birth:"1968-03-08 1:00:00 pm"
      }
   \]
}

If we use the template like following:

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth: <<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

  

The results will be produced for both pieces of JSON data like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format becomes possible, because text values of Birth properties are automatically converted to DateTime?.{{< /alert >}}

### Treating the objects at top level

If a top-level JSON element represents an object, a *JsonDataSource* instance should be treated in template documents in the same way as if it was a *DataRow* instance (see "[Working with *DataRow* and *DataRowView* Objects](https://docs.groupdocs.com/display/assemblynet/Template+Syntax+-+Part+1+of+2#TemplateSyntax-Part1of2-DataTableandDataViewObjects)" for more information). If a top-level JSON object has a single property that is also an object, then this nested object is accessed by the assembler instead.

Suppose we have Json data like:

{ 
   Name:"John Doe",
   Age:30,
   Birth:"1989-04-01 4:00:00 pm",
   Child:\[ 
      "Ann Doe",
      "Charles Doe"
   \]
}

or alternatively like:

{ 
   Person:{ 
      Name:"John Doe",
      Age:30,
      Birth:"1989-04-01 4:00:00 pm",
      Child:\[ 
         "Ann Doe",
         "Charles Doe"
      \]
   }
}

If we use the template like following:

Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth:
<<\[Birth\]:"dd.MM.yyyy">>
Children:
<<foreach \[in Child\]>><<\[Child\_Text\]>>
<</foreach>>

The results will be produced for both pieces of JSON data like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Children:
Ann Doe
Charles Doe

{{< alert style="warning" >}}To reference a JSON object property that is an array of simple-type values, the name of the property (for example, "Child") should be used in a template document, whereas the same name with the "_Text" suffix (for example, "Child_Text") should be used to reference the value of an item of this array.{{< /alert >}}

### The complete example

The following example sums up typical scenarios involving nested JSON objects and arrays.

#### JSON

\[ 
   { 
      Name:"John Smith",
      Contract:\[ 
         { 
            Client:{ 
               Name:"A Company"
            },
            Price:1200000
         },
         { 
            Client:{ 
               Name:"B Ltd."
            },
            Price:750000
         },
         { 
            Client:{ 
               Name:"C & D"
            },
            Price:350000
         }
      \]
   },
   { 
      Name:"Tony Anderson",
      Contract:\[ 
         { 
            Client:{ 
               Name:"E Corp."
            },
            Price:650000
         },
         { 
            Client:{ 
               Name:"F & Partners"
            },
            Price:550000
         }
      \]
   },
   { 
      Name:"July James",
      Contract:\[ 
         { 
            Client:{ 
               Name:"G & Co."
            },
            Price:350000
         },
         { 
            Client:{ 
               Name:"H Group"
            },
            Price:250000
         },
         { 
            Client:{ 
               Name:"I & Sons"
            },
            Price:100000
         },
         { 
            Client:{ 
               Name:"J Ent."
            },
            Price:100000
         }
      \]
   }
\]

#### Template document

<<foreach \[in managers\]>>Manager: <<\[Name\]>>
Contracts:
<<foreach \[in Contract\]>>- <<\[Client.Name\]>> ($<<\[Price\]>>)
<</foreach>>
<</foreach>>

#### Source code

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-simplejsonds_demo_19-10-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-simplejsonds_demo_19-10-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-simplejsonds_demo_19-10-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source document template</span></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-simplejsonds_demo_19-10-cs-LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Templates/SimpleDatasetDemo.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-simplejsonds_demo_19-10-cs-LC4" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-simplejsonds_demo_19-10-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination for reports</span></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-simplejsonds_demo_19-10-cs-LC6" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/SimpleJsonDSDemo Out.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-simplejsonds_demo_19-10-cs-LC7" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-simplejsonds_demo_19-10-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination Markdown reports</span></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-simplejsonds_demo_19-10-cs-LC9" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDataSource</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>JSON DataSource/ManagerData.json<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-simplejsonds_demo_19-10-cs-LC10" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-simplejsonds_demo_19-10-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>initialize data source</span></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-simplejsonds_demo_19-10-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-en">JsonDataSource</span> <span class="pl-smi">dataSource</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">JsonDataSource</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetDataSourceDocument</span>(<span class="pl-smi">strDataSource</span>));</td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-simplejsonds_demo_19-10-cs-LC13" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-simplejsonds_demo_19-10-cs-LC14" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Assemble document</span></td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-simplejsonds_demo_19-10-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-simplejsonds_demo_19-10-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-simplejsonds_demo_19-10-cs-LC16" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>), <span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">dataSource</span>, <span class="pl-s"><span class="pl-pds">"</span>managers<span class="pl-pds">"</span></span>));</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/330892e76bb1958b9199e3d223962e27/raw/3c5c9c4597332ee0fba2ff28803e76bcbfc2b327/SimpleJsonDS_Demo_19.10.cs) [SimpleJsonDS\_Demo\_19.10.cs](https://gist.github.com/GroupDocsGists/330892e76bb1958b9199e3d223962e27#file-simplejsonds_demo_19-10-cs) hosted with ❤ by [GitHub](https://github.com)

  

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

*   [ManagerData.json](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/JSON%20DataSource/ManagerData.json?raw=true)

#### Template

*   [Demo Template.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)
