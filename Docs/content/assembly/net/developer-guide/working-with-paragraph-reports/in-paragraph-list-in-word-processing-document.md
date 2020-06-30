---
id: in-paragraph-list-in-word-processing-document
url: assembly/net/in-paragraph-list-in-word-processing-document
title: In-Paragraph List in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-ParagraphList report in Word Processing Document format based on the use case: Working with a Business Case.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## In-Paragraph List in Microsoft Word Document

### Creating a In-Paragraph List

Practicing the following steps you can create In-Paragraph List Template in MS Word 2013.

1.  Write a sentence, for example "We provide support for the following products:".
2.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   A descriptive or informative line like "We provide support for the following products:".
*   Show all the products along with the above sentence.
*   Report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[IndexOf() != 0 ? ", " : ""]>>
<<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download In-Paragraph List Template

Please download the sample In-Paragraph List document we created in this article:

*   [In-Paragraph List.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Paragraph%20List.docx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [In-Paragraph List\_XML.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Paragraph%20List_XML.docx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 7775ae4897672251532b3d83e751747f >}}



#### Database Entities

{{< gist GroupDocsGists 68e2eb6091b3ee906eba1f90e5622ec3 >}}



#### Using DataSet

{{< gist GroupDocsGists 6bdec421567135d03c820817338b2c1a >}}



#### Using XML DataSource

{{< gist GroupDocsGists c40e3a2f3a7e262bc7b538293a0f9d36 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 9164646f73917ebe7e06afa78c28e333 >}}



## In-Paragraph List in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [Wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'In-Paragraph List' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-paragraph-reports/in-paragraph-list-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Paragraph List\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Paragraph%20List_OpenDocument.odt?raw=true) (use same template for database, dataset and JSON examples)
*   [In-Paragraph List\_XML\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Paragraph%20List_XML_OpenDocument.odt?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists ad17c8a96d43e55fd31a2e786eff308d >}}



#### Database Entities

{{< gist GroupDocsGists 41977a1de07bf0a29e41d83de181bb9b >}}



#### Using DataSet

{{< gist GroupDocsGists aa9772303065899b4860b8a951bbe3ad >}}



#### Using XML DataSource

{{< gist GroupDocsGists 70e8a390639f9334e01720353ca2cf35 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 9a4c83a74acee10a435682fc887af522 >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
