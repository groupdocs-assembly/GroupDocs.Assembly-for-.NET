---
id: numbered-list-in-presentation-document
url: assembly/net/numbered-list-in-presentation-document
title: Numbered List in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Numbered List report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Numbered List in Microsoft PowerPoint Document

### Creating a Numbered List

Practising the following steps you can create Numbered List Template in MS PowerPoint 2013.

1.  Add a new presentation slide.
2.  Write a sentence like "We provide support for the following products:".
3.  Start numbered list.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in numbered list.
*   The report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
1.	<<foreach [in products]>><<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

Download Numbered List Template

Please download the sample Numbered List document we created in this article:

*   [Numbered List.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Numbered%20List.pptx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Numbered List\_XML.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Numbered%20List_XML.pptx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 9938d1ac7bfd4055d7d9e320740d5b7c >}}



#### Database Entities

{{< gist GroupDocsGists 6c0521527ce3db2a2e9dc7ff226afec9 >}}



#### Using DataSet

{{< gist GroupDocsGists 907febef1b581c1581f030e3544b07d3 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 714f7c87e58432e958b3c9dba775ca10 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 2ed14faca3d0cad68134c45c0e4148ea >}}



## Numbered List in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [Wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Numbered List' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-list-reports-numbered/numbered-list-in-presentation-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Presentation" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Numbered List\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Numbered%20List_OpenDocument.odp?raw=true) (use same template for database, dataset and JSON examples)
*   [Numbered List\_XML\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Numbered%20List_XML_OpenDocument.odp?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists ffa616e4e2add07dbe09e961622a9bb3 >}}



#### Database Entities

{{< gist GroupDocsGists 21f96063d1786eac6c824d0997498f91 >}}



#### Using DataSet

{{< gist GroupDocsGists f21b85a611d8c90fd2209dfbd7fc7ef3 >}}



#### Using XML DataSource

{{< gist GroupDocsGists c63cf0d328d422573d8969157bdf1ce9 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 350bf392a2baa8220b97363f9bcff887 >}}



### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
