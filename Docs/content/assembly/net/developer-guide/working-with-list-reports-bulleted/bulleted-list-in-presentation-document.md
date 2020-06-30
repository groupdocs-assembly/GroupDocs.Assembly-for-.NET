---
id: bulleted-list-in-presentation-document
url: assembly/net/bulleted-list-in-presentation-document
title: Bulleted List in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bulleted List report in Presentation Document format based on the use case: Working with a Business Case.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Bulleted List in Microsoft PowerPoint Document

### Creating a Bulleted List

Practising the following steps you can insert Bulleted List in MS PowerPoint 2013.

1.  Add a new presentation slide.
2.  Add a bullet list at the place where you want it.
3.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share the information of the products with the following key requirements:

*   A report must show all the products in a bulleted list format.
*   A report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

### Download Bulleted List Template

Please download the sample Bulleted List document we created in this article:

*   [Bulleted List.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Bulleted%20List.pptx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Bulleted List\_XML.pptx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Bulleted%20List_XML.pptx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists bed51427de15e362dcaa0c16dd6e8e94 >}}



#### Database Entities

{{< gist GroupDocsGists bfe5bed009bd025227e3bc6fb68f9e10 >}}



#### Using DataSet

{{< gist GroupDocsGists be387a246bd2af94d9b38bfc5bb6986b >}}



#### Using XML DataSource

{{< gist GroupDocsGists 31cc8555ad883e9427094d06a484f82d >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 7b364145e5fe1c8279706088c0459273 >}}



## Bulleted List in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Bulleted List' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-list-reports-bulleted/bulleted-list-in-presentation-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Presentation" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Bulleted List\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Bulleted%20List_OpenDocument.odp?raw=true) (use same template for database, dataset and JSON examples)
*   [Bulleted List\_XML\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Bulleted%20List_XML_OpenDocument.odp?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 2c0921883aa5cf90814f6750177a746f >}}



#### Database Entities

{{< gist GroupDocsGists fe151d8f3fb60ae08a7349688eaf665e >}}



#### Using DataSet

{{< gist GroupDocsGists f796a16bce9c2a864526d434c0dcf896 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 994f15754722fd2245978b7023adf519 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists aad2aa589ce8873dc9d1ae4a7cba6c88 >}}



### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
