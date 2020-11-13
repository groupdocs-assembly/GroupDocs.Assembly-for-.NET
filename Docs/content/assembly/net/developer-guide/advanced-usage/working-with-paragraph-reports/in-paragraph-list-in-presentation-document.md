---
id: in-paragraph-list-in-presentation-document
url: assembly/net/in-paragraph-list-in-presentation-document
title: In-Paragraph List in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Paragraph List report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}

## In-Paragraph List in Microsoft PowerPoint Document

### Creating a In-Paragraph List

Practicing the following steps you can create In-Paragraph List Template in MS PowerPoint 2013.

1.  Add a new presentation slide.
2.  Write a sentence, for example "We provide support for the following products:".
3.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   A descriptive or informative line like "We provide support for the following products:".
*   Show all the products along with the above sentence.
*   Report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[IndexOf() != 0 ? ", " : ""]>>
<<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download In-Paragraph List Template

Please download the sample In-Paragraph List document we created in this article:

*   [In-Paragraph List.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Paragraph%20List.pptx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [In-Paragraph List\_XML.pptx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Paragraph%20List_XML.pptx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs bd5557c6afce363c91b4 >}}

#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 7af1f7c4d8657f446fbb >}}

#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 5eece4a7628ce469d578 >}}

#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 93708a50854ca5c5bae8 >}}

#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 213cf34e8727ab1d224a >}}

## In-Paragraph List in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'In-Paragraph List' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Presentation" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Paragraph List\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Paragraph%20List_OpenDocument.odp?raw=true) (use same template for database, dataset and JSON examples)
*   [In-Paragraph List\_XML\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Paragraph%20List_XML_OpenDocument.odp?raw=true)

### Generating the Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 67b2109c81d174ecea7c >}}

#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 1ce821c079b3f54c1272 >}}

#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 464b40a30fc1fee5cd38 >}}

#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 8123d5ed9e9d796def53 >}}

#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 0342489311290a4fa12c >}}

### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
