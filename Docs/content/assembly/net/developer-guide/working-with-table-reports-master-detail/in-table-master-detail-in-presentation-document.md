---
id: in-table-master-detail-in-presentation-document
url: assembly/net/in-table-master-detail-in-presentation-document
title: In-Table Master-Detail in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate an In-Table Master-Detail report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## In-Table Master-Detail in Microsoft PowerPoint Document

### Creating a In-Table Master-Detail

Practising the following steps you can create In-Table Master-Detail Template in MS PowerPoint 2013.

1.  Add a new presentation slide.
2.  Click the document where you want to add the table.
3.  Press "Insert" tab to insert the table.
4.  Insert a 2x4 table.
5.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent customers' total orders prices and price of each product separately with the following key requirements:

*   The report must show each customer along with his total orders prices.
*   It must also show each individual product ordered by the customers.
*   It must show the sum of the order prices.
*   It must represent all the information in tabular form.
*   The report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 457pt;"><tbody><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 3pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span></p></td><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 3pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">Order Price</span></p></td></tr><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 3pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">foreach</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;"> [in customers]&gt;&gt;&lt;&lt;[</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">CustomerName</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(210, 222, 239); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 3pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Order.Sum</span><span style="font-family: Calibri; font-size: 11pt;">(</span></p><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-style: italic;">&lt;&lt;</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-style: italic;">foreach</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-style: italic;"> [in Order]&gt;&gt; &lt;&lt;[</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-style: italic;">Product.ProductName</span><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-style: italic;">]&gt;&gt;</span></p></td><td style="background-color: rgb(234, 239, 247); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Price]&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;&lt;&lt;</span></p><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(91, 155, 213); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 267.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(255, 255, 255); font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="background-color: rgb(210, 222, 239); border-bottom-color: rgb(255, 255, 255); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(255, 255, 255); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(255, 255, 255); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(255, 255, 255); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; padding-top: 0.25pt; vertical-align: top; width: 167.2pt;"><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Sum(</span></p><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">m =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">m.Order.Sum</span><span style="font-family: Calibri; font-size: 11pt;">(</span></p><p style="font-size: 11pt; line-height: 107%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">))]&gt;&gt;</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download In-Table Master-Detail Template

Please download the sample In-Table Master-Detail document we created in this article:

*   [In-Table Master-Detail.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail.pptx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table Master-Detail\_DB.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_DB.pptx?raw=true) (Template for DataBase examples)
*   [In-Table Master-Detail\_DT.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_DT.pptx?raw=true) (Template for DataSet examples)
*   [In-Table Master-Detail\_XML.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_XML.pptx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists c9c8b151dd0e46be909be454a6726049 >}}



#### Database Entities

{{< gist GroupDocsGists de3833fac1caad796cf165a8bae6456b >}}



#### Using DataSet

{{< gist GroupDocsGists 82d6372ad526fe33d0993be1300bbb43 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 1f5520cbc9f6793fa423eda6e936b77a >}}



#### Using JSON DataSource

{{< gist GroupDocsGists facbd1591cf19cc15961a3e40a4a50b5 >}}



## In-Table Master-Detail in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table Master-Detail' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-table-reports-master-detail/in-table-master-detail-in-presentation-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Presentation" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table Master-Detail\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_OpenDocument.odp?raw=true) (use same template for JSON examples)
*   [In-Table Master-Detail\_DB\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_DB_OpenDocument.odp?raw=true)
*   [In-Table Master-Detail\_DT\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_DT_OpenDocument.odp?raw=true)
*   [In-Table Master-Detail\_XML\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/In-Table%20Master-Detail_XML_OpenDocument.odp?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists dbae5688546ba5feaa8326b2cd74ab9c >}}



#### Database Entities

{{< gist GroupDocsGists d843d458b5c9c8c85b6f2d8d35bfd569 >}}



#### Using DataSet

{{< gist GroupDocsGists 4d95131b1ff98587b3482abe85ed78b5 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 3e380308021b5b5306244a30aabab2d9 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists e41f5bc6342e8b91b48a1a426918b2fe >}}



### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
