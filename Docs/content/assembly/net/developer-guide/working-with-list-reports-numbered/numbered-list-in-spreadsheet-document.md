---
id: numbered-list-in-spreadsheet-document
url: assembly/net/numbered-list-in-spreadsheet-document
title: Numbered List in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Numbered List report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Numbered List in Microsoft Excel Document

### Creating a Numbered List

Practising the following steps you can create Numbered List Template in MS Excel 2013.  
You can create two kinds of Numbered List in Microsoft Excel.

*   In-cell List
*   Multiple-cell List

Apply the following steps to build the In-Cell template

1.  Add a new Workbook.
2.  Write a sentence, for example "We provide support for the following products:" in a single cell and add syntax there.
3.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in numbered list.
*   The report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

**In-cell List**

```csharp
"We provide support for the following products:<<foreach [in products]>>
<<[NumberOf()]>>.         <<[ProductName]>><</foreach>>"
```

**Multiple-cell List**

```csharp
<<foreach [in products]>><<[NumberOf()]>>.         <<[ProductName]>>
```

and close foreach in next column,

```csharp
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

Download Numbered List Template

Please download the sample Numbered List document we created in this article:

*   [Numbered List.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Numbered%20List.xlsx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Numbered List\_XML.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Numbered%20List_XML.xlsx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 2fbc89dd5cea0ea6b55adb5aeafdf801 >}}



#### Database Entities

{{< gist GroupDocsGists 2dd32d44fd6fd9128b2ea32620dddf62 >}}



#### Using DataSet

{{< gist GroupDocsGists a1c29ccafcec8efd2d58b4a7d6bbcc0d >}}



#### Using XML DataSource

{{< gist GroupDocsGists d00d843a5c02f5477c1f90f750361e67 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 60825204f717ea6c24b577c8b8061ca9 >}}



## Numbered List in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [Wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Numbered List' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-list-reports-numbered/numbered-list-in-spreadsheet-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Numbered List\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Numbered%20List_OpenDocument.ods?raw=true) (use same template for database, dataset and JSON examples)
*   [Numbered List\_XML\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Numbered%20List_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists f7f8f23a66c5efcec325927c6825d419 >}}



#### Database Entities

{{< gist GroupDocsGists b204172f799f7e3371eb39f54d67bc66 >}}



#### Using DataSet

{{< gist GroupDocsGists 10a774241943e46d9dde65c6b03a38e9 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 4d8b6adb9bd06a0644ffaed53bfe88d6 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists fa5af16289e3dd4a5fb2892163ed758a >}}



### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
