---
id: common-master-detail-image-in-presentation-document
url: assembly/net/common-master-detail-image-in-presentation-document
title: Common Master-Detail Image in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}

## Common Master-Detail in Microsoft PowerPoint Document

### Creating a Common Master-Detail

Please follow below steps to create Common Master-Detail Template in MS PowerPoint 2013:

1.  Create a new Presentation.
2.  Insert two shapes, one for holding image and other for holding other details.  
  
## Save the document

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   Report must show customers' picture and name.
*   It must associate the customers with their products.
*   Report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

NOTE: There is no way to use images inside a foreach tag in PowerPoint.

<<foreach \[in customers\]>>**Customer: <<\[CustomerName\]>>**

**Products: <<foreach \[in Order\]>><<\[IndexOf() != 0 ? ", " : ""\]>><<\[Product.ProductName\]>><</foreach\>>**

<</foreach\>>

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Common Master-Detail Template

Please download the sample Common Master-Detail document we created in this article:

*   [Common Master-Detail.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Common%20Master-Detail.pptx?raw=true) (Template for CustomObject and JSON examples)
*   [Common Master-Detail\_DB.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Common%20Master-Detail_DB.pptx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs cc4d171e25ad9761120f >}}

#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 9af73f2cc66a6477331c >}}

#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs e8dd676f698f6ddc5b85 >}}

#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs ec5c65697d59406bdece >}}

#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 155eb16766da06a0ea9d >}}

## Common Master-Detail in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Common Master-Detail' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Common Master-Detail\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Common%20Master-Detail_OpenDocument.odp?raw=true) (use same template for JSON examples)
*   [Common Master-Detail\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Common%20Master-Detail_DB_OpenDocument.odp?raw=true) (use same template for database and dataset and XML examples)

### Generating the report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 56f152d91f84cefef8d7 >}}

#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 680319f37bc74988a684 >}}

#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 5392574bdb04e851863e >}}

#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 8e2f524e84ca847fcfd4 >}}

#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs f59b2aa5a757bf90336a >}}

### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
