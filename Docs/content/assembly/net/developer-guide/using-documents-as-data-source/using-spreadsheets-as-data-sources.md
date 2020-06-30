---
id: using-spreadsheets-as-data-sources
url: assembly/net/using-spreadsheets-as-data-sources
title: Using Spreadsheets as Data Sources
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 16.12.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Using Spreadsheets as Data Sources

Following classes are added in [GroupDocs.Assembly.Data](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/) namespace:

*   [DocumentTable](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttable)
*   [DocumentTableColumn](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablecolumn)
*   [DocumentTableColumnCollection](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttablecolumncollection)
*   [DocumentTableOptions](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/documenttableoptions)

## The Recipe

*   Define template and output report documents
*   Assemble a document using the external document table as a data source

### Download

#### Data Source Document

*   [Contracts Data.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Excel%20DataSource/Contracts%20Data.xlsx?raw=true)

#### Template

*   [Using Spreadsheet as Table of Data.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)

## The Code

{{< gist GroupDocsGists 2cedc6e4a5c1e3f79c370104486124b4 >}}


