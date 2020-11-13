---
id: multicolored-numbered-list-in-presentation-document
url: assembly/net/multicolored-numbered-list-in-presentation-document
title: Multicolored Numbered List in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Multicolored Numbered List report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Multicolored Numbered List in Microsoft PowerPoint Document

### Creating a Multicolored Numbered List

Practising the following steps you can create Multicolored Numbered List Template in MS PowerPoint 2013.

1.  In your document, write a sentence like "We provide support for the following products:".
2.  Start numbered list.
3.  Go to the "Design" tab and select color to make it colored list.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in numbered list.
*   It must highlight the products.
*   The report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 457pt;"><tbody><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 3pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span></p></td><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 3pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order Price</span></p></td></tr><tr><td style="background-color: rgb(244, 177, 131); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 3pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;foreach [in orders]&gt;&gt;&lt;&lt;if [Price &gt;= 400]&gt;&gt;&lt;&lt;[Customer.CustomerName]&gt;&gt;</span></p></td><td style="background-color: rgb(244, 177, 131); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 3pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Price]&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;[Customer.CustomerName]&gt;&gt;</span></p></td><td style="background-color: rgb(234, 239, 247); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Price]&gt;&gt;&lt;&lt;/if&gt;&gt;&lt;&lt;/foreach&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="background-color: rgb(210, 222, 239); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Sum(c =&gt; c.Price)]&gt;&gt;</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

Download Multicolored Numbered List Template

Please download the sample Multicolored Numbered List document we created in this article:

*   [Multicolored Numbered List.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Multicolored%20Numbered%20List.pptx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Multicolored Numbered List\_XML.pptx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Multicolored%20Numbered%20List_XML.pptx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 5e5f0065e670db7d7467c24bc9ab20b1 >}}

#### Database Entities

{{< gist GroupDocsGists ec5aa9bf0aef124323d6be3cf64103dc >}}

#### Using DataSet

{{< gist GroupDocsGists 4d6374c206d0c5542eeac637f1f6542c >}}

#### Using XML DataSource

{{< gist GroupDocsGists 0237cc5e3df88002b22ac8580bcbe039 >}}

#### Using JSON DataSource

{{< gist GroupDocsGists bb8683af4ca31227bb36954b15b21299 >}}

## Multicolored Numbered List in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Multicolored Numbered List' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Presentation" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Multicolored Numbered List\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Multicolored%20Numbered%20List_OpenDocument.odp?raw=true) (use same template for database, dataset and JSON examples)
*   [Multicolored Numbered List\_XML\_OpenDocument.odp](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Multicolored%20Numbered%20List_XML_OpenDocument.odp?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 45de4042f2e5280dd590dcd944663d49 >}}

#### Database Entities

{{< gist GroupDocsGists efa64394b5d02a0d4c2371933424fb78 >}}

#### Using DataSet

{{< gist GroupDocsGists 1bf5a2ee72180eeb1855c15d9926e840 >}}

#### Using XML DataSource

{{< gist GroupDocsGists a0aa38544a9141103f8a3bedfc69a07f >}}

#### Using JSON DataSource

{{< gist GroupDocsGists fa64bc092425100112d358209f031742 >}}

### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
