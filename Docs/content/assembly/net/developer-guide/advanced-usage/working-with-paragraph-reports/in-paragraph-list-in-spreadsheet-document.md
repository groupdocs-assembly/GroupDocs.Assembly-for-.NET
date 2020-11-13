---
id: in-paragraph-list-in-spreadsheet-document
url: assembly/net/in-paragraph-list-in-spreadsheet-document
title: In-Paragraph List in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Paragraph List report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}

## In-Paragraph List in Microsoft Excel Document

### Creating a In-Paragraph List

Practicing the following steps you can create In-Paragraph List Template in MS Excel 2013.

1.  Add a new Workbook.
2.  Write a sentence, for example "We provide support for the following products:" in a single column.
3.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe about the services you are providing with the following key requirements:

*   A descriptive or informative line like "We provide support for the following products:".
*   Show all the products along with the above sentence.
*   Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[IndexOf() != 0 ? ", " : ""]>>
<<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download In-Paragraph List Template

Please download the sample In-Paragraph List document we created in this article:

*   [In-Paragraph List.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Paragraph%20List.xlsx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [In-Paragraph List\_XML.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Paragraph%20List_XML.xlsx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 6b9866627791c4ff0a3c >}}

#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 6223b2987be48a027f99 >}}

#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 25c0b61ebda5c39421f2 >}}

#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 6ed5df46b0ef742488ac >}}

#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs dac0a613622ec857902d >}}

## In-Paragraph List in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'In-Paragraph List' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Paragraph List\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Paragraph%20List_OpenDocument.ods?raw=true) (use same template for database, dataset and JSON examples)
*   [In-Paragraph List\_XML\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Paragraph%20List_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 6445b2968162214aaf91 >}}

#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 9fa306dad3920785fcc8 >}}

#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 6eef45055feb79ed238e >}}

#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 614f40dafb84838a9452 >}}

#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 66fc77bc84df4f1ec7fc >}}

### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
