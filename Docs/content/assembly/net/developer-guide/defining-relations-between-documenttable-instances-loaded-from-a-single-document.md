---
id: defining-relations-between-documenttable-instances-loaded-from-a-single-document
url: assembly/net/defining-relations-between-documenttable-instances-loaded-from-a-single-document
title: Defining Relations Between DocumentTable Instances Loaded from a Single Document
weight: 34
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Define Relations Between Document Table Instances

API provides ability to define relations between [DocumentTable](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttable) instances loaded from a single document. The following classes of the [GroupDocs.Assembly.Data](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/) namespace have been added:

*   [DocumentTableRelation](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablerelation)
*   [DocumentTableRelationCollection](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablerelationcollection)

Moreover, the **Relations** property of the [GroupDocs.Assembly.Data.DocumentTableSet](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttableset) class has been added.

### Data Source Document

*   [Related Tables Data.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Excel%20DataSource/Related%20Tables%20Data.xlsx?raw=true)

### Template

*   [Using Document Table Relations.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Document%20Table%20Relations.docx?raw=true)

## Using Document Table Relations

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist GroupDocsGists ce6189b20fa197fee3b15833b26127fc >}}

## ColumnNameExtractingDocumentTableLoadHandler

{{< gist GroupDocsGists a855621e468c2c627534255dbcb7657c >}}


