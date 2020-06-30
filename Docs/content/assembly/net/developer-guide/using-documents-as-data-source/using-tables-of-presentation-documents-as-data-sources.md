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

{{< gist GroupDocsGists b69cdd70c2f91362b2c6fdddd7bea8a5 >}}


