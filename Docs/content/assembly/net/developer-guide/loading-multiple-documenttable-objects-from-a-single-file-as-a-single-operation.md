---
id: loading-multiple-documenttable-objects-from-a-single-file-as-a-single-operation
url: assembly/net/loading-multiple-documenttable-objects-from-a-single-file-as-a-single-operation
title: Loading Multiple DocumentTable Objects from a Single File as a Single Operation
weight: 33
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Loading Multiple DocumentTable Objects

GroupDocs.Assembly for .NET API provides the ability to load multiple DocumentTable objects from a single file as a single operation. Following classes and interfaces of the [GroupDocs.Assembly.Data](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/) namespace have been added:

*   [DocumentTableSet](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttableset)
*   [DocumentTableCollection](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablecollection)
*   [IDocumentTableLoadHandler](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/idocumenttableloadhandler)
*   [DocumentTableLoadArgs.](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttableloadargs)

Moreover, following properties of the [GroupDocs.Assembly.Data.DocumentTable](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttable) class have been added:

*   Name
*   IndexInDocument

### Data Source Document

*   [Multiple Tables Data.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Word%20DataSource/Multiple%20Tables%20Data.docx?raw=true)

### Template

*   [Using Document Table Set as Data Source.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Using%20Document%20Table%20Set%20as%20Data%20Source.pptx?raw=true)

## Loading DocumentTableSet using Default Options

{{< gist GroupDocsGists a6f71eb9b5b06dc1c9038fd25965ade9 >}}

## Loading DocumentTableSet using Custom Options

{{< gist GroupDocsGists b9765c9a2d1ab7a67c4215eec42306fd >}}

## Using DocumentTableSet as Data Source

{{< gist GroupDocsGists 5a753fa52864e3126a697ac0f0d96b15 >}}

## CustomDocumentTableLoadHandler

{{< gist GroupDocsGists 6c1d6116bd499686ad3cec225423b3ce >}}

## ColumnNameExtractingDocumentTableLoadHandler

{{< gist GroupDocsGists a855621e468c2c627534255dbcb7657c >}}
