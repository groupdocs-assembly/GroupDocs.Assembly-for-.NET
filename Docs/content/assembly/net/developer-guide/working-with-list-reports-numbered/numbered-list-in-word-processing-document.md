---
id: numbered-list-in-word-processing-document
url: assembly/net/numbered-list-in-word-processing-document
title: Numbered List in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Numbered List report in Word Processing Document format based on the use case: Working with a Business Case.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Numbered List in Microsoft Word Document

### Creating a Numbered List

Practising the following steps you can create Numbered List Template in MS Word 2013.

1.  In your document, write a sentence like "We provide support for the following products:".
2.  Start numbered list.
3.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in the numbered list.
*   The report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
1.	<<foreach [in products]>><<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

Download Numbered List Template

Please download the sample Numbered List document we created in this article:

*   [Numbered List.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Numbered%20List.docx?raw=true) (Template for DataBase, DataSet and JSON Data examples)
*   [Numbered List\_XML.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Numbered%20List_XML.docx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 09a797b200c1540363f826971a0d77d3 >}}



#### Database Entities

{{< gist GroupDocsGists 96b6ec9337b6da478cca8dff8ddb6108 >}}



#### Using DataSet

{{< gist GroupDocsGists d3018f38635db03060cbf8f61f1050a9 >}}



#### Using XML DataSource

{{< gist GroupDocsGists ab228e8733227e4f70082bf65400b6d8 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists fd6cf2b09033f95ef8adc3c19004433f >}}



## Numbered List in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Numbered List' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-list-reports-numbered/numbered-list-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Numbered List\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Numbered%20List_OpenDocument.odt?raw=true) (use same template for database, dataset and JSON examples)
*   [Numbered List\_XML\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Numbered%20List_XML_OpenDocument.odt?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists cae108ae98bff77def2058d3712429c6 >}}



#### Database Entities

{{< gist GroupDocsGists 0ac8f49a57de8ea4d9d87e499a594808 >}}



#### Using DataSet

{{< gist GroupDocsGists 1e40425f2ab1ecf4d97bbdc48bd14940 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 13b78920844bf08e8609722c98edd401 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists c8ddee4c4a1fc5aad412eef350b830ef >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.

## Numbering Restart in Nested Numbered List 

{{< alert style="info" >}}This feature is supported by version 19.7 or greater.{{< /alert >}}

  
The GroupDocs.Assembly engine allows restart list numbering within your documents dynamically using *<<restartNum>>* tags. In particular, this feature is useful when working with a nested numbered list within a data band.  

Assume that we are picking *Order* and *Service* classes as defined in the following *Custom Objects* of our business use case.

{{< gist GroupDocsGists cd8768711ff88414192c3f50a9b7c918 >}}



Given that orders is an enumeration of Order instances, you could try to use the following template to output information on several orders in one document.

```csharp
<<foreach [in orders]>><<[Customer.CustomerName]>> (<<[Customer.CustomerContactNumber]>>)
1.	<<foreach [in Services]>><<[ServiceName]>>
<</foreach>><</foreach>>
```

The generated report will look as follows:

Jane Doe (+9211874)
	1.	Regular Cleaning
	2.	Oven Cleaning
John Smith (+458789)
	3.	Regular Cleaning
	4.	Oven Cleaning
	5.	Carpet Cleaning
John Smith (+458789)
	6.	Regular Cleaning
	7.	Carpet Cleaning
John Smith (+458789)
	8.	Oven Cleaning

However, there would be a single numbered list across all orders, which is not applicable for this scenario. Hence, you can make the list numbering to restart for every order by putting a *<<restartNum>>* tag into your template before a corresponding *<<foreach>>* tag as follows:

```csharp
<<foreach [in orders]>><<[ Customer.CustomerName]>> (<<[Customer.CustomerContactNumber]>>)
1.	<<restartNum>><<foreach [in Services]>><<[ServiceName]>>
<</foreach>><</foreach>>
```

 Then, the generated report will look as follows:

Jane Doe (+9211874)
	1.	Regular Cleaning
	2.	Oven Cleaning
John Smith (+458789)
	1.	Regular Cleaning
	2.	Oven Cleaning
	3.	Carpet Cleaning
John Smith (+458789)
	1.	Regular Cleaning
	2.	Carpet Cleaning
John Smith (+458789)
	1.	Oven Cleaning

Download Numbered List Template

Please download the sample Numbered List template we created in this article:

*   [Numbered List\_RestartNum.docx](attachments/34439255/85426181.docx) (Template for DataBase, DataSet and JSON Data examples)
