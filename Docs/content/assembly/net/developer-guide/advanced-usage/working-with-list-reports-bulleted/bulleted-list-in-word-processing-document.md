---
id: bulleted-list-in-word-processing-document
url: assembly/net/bulleted-list-in-word-processing-document
title: Bulleted List in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bulleted List report in Word Processing Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Bulleted List in Microsoft Word Document

### Creating a Bulleted List

Practising the following steps you can insert Bulleted List in MS Word 2013.

1.  Add a bullet list at the place where you want it.
2.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share the information of the products with the following key requirements:

*   Report must show all the products in a bulleted list format.
*   Report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

**Template Syntax**

```csharp
We provide support for the following products:
. <<foreach [in products]>><<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Bulleted List Template

Please download the sample Bulleted List document we created in this article:

*   [Bulleted List.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Bulleted%20List.docx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Bulleted List\_XML.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Bulleted%20List_XML.docx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 9b55cff16b005121e40cc9080920821a >}}



#### Database Entities

{{< gist GroupDocsGists 28411995912ff7684a0dce651dcb5f09 >}}



#### Using DataSet

{{< gist GroupDocsGists 73097d50587221a303300a8355f4608f >}}



#### Using XML DataSource

{{< gist GroupDocsGists 17d59b45b1d6056076533e12f2b4008e >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 2a3c9f8c8ed6f9ce9e1e686ff7245078 >}}



## Bulleted List in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Bulleted List' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Bulleted List\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Bulleted%20List_OpenDocument.odt?raw=true) (use same template for database, dataset and JSON examples)
*   [Bulleted List\_XML\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Bulleted%20List_XML_OpenDocument.odt?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 97560bb8f4b8d87296065dbc1297b21e >}}



#### Database Entities

{{< gist GroupDocsGists a29cba37db08cbb72fc17e55d8af9597 >}}



#### Using DataSet

{{< gist GroupDocsGists 96f83f5c5af219b723240fb4aa1ffd50 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 9c2371892822f0ac46cf7ad61e16fa14 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 9b056e349187d03f4c6d8f1aab80c6a0 >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
