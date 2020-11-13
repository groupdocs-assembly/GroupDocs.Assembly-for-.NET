---
id: single-row-image-in-presentation-document
url: assembly/net/single-row-image-in-presentation-document
title: Single Row Image in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Single Row Image report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}

## Single Row in Microsoft PowerPoint Document

### Creating a Single Row

Please follow below steps to create Single Row Image in MS PowerPoint 2013:

1.  Create a new presentation slide.
2.  Insert the desired shape to display the image in it.
3.  Go to Insert Tab and select shape by clicking on Shape Icon.
4.  Write "Name" and "Contact Number" to show the name and contact info of the customer.
5.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent information of first single customer with the following key requirements:

*   Report must show image of the customer.
*   It must show Name and Contact Number of the customer.
*   Report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

{{< alert style="warning" >}}At the moment, GroupDocs.Assembly does not support images in PowerPoint Presentation.{{< /alert >}}

<table class="iksweb">
	<tbody>
		<tr>
			<td rowspan="2" style="font-size: 44pt">Click to<br> add title</td>
			<td>Name:</td>
			<td><<[customer.CustomerName]>></td>
		</tr>
		<tr>
			<td>Contact Details:</td>
			<td><<[customer. CustomerContactNumber]>></td>
		</tr>
	</tbody>
</table>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Single Row Template

Please download the sample Single Row document we created in this article:

*   [Single Row.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Single%20Row.pptx?raw=true)

{{< alert style="warning" >}}Use this template for DB, Dataset, JSON and XML examples also.{{< /alert >}}

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 81fa18ead6c31fc0ce67865d45fe721a >}}

#### Database Entities

{{< gist GroupDocsGists af2e23cb9e089ac183b646d878cd32ad >}}

#### Using DataSet

{{< gist GroupDocsGists 41f796d33332cbce2a6be8f44b73c5aa >}}

#### Using XML DataSource

{{< gist GroupDocsGists 5160c2dedcf888126c41ed2a7318457e >}}

#### Using JSON DataSource

{{< gist GroupDocsGists 1c5b5322890dca734b2b0eeb464a67e4 >}}

## Single Row in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Single Row' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Single Row\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Single%20Row_OpenDocument.odp?raw=true) (use same template for database, dataset, JSON and XML examples)

### Generating the report

#### Custom Objects

{{< gist GroupDocsGists f44876d4e43c66fb3b536d56847aa8f4 >}}

#### Database Entities

{{< gist GroupDocsGists 317a91c44b5618a36e29cc3d8462ae12 >}}

#### Using DataSet

{{< gist GroupDocsGists 1069f4bf98e493f1b0937874dbf62b32 >}}

#### Using XML DataSource

{{< gist GroupDocsGists 4d22e44311fd0efdaa08a0044db5e586 >}}

#### Using JSON DataSource

{{< gist GroupDocsGists fd80462fa882ae524a2b4295777afbcf >}}

### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
