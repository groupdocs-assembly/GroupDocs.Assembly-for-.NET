---
id: common-list-image-in-presentation-document
url: assembly/net/common-list-image-in-presentation-document
title: Common List Image in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common List Image report in Presentation Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Common List in Microsoft PowerPoint Document

### Creating a Common List Image

Please follow below steps to create Common List Template in MS PowerPoint 2013:

1.  Create a new presentation slide.
2.  Add two shapes to hold pictures and name.
3.  Save your Document.

### Reporting Requirement

As a report developer, you are required to represent the information of the customers with the following key requirements:

*   A report must show multiple customers' picture and name.
*   A report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

{{< alert style="warning" >}}There is no way to use an image inside a foreach tag in Microsoft PowerPoint.{{< /alert >}}

```csharp
<<foreach [in customers]>><<[CustomerName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Template

Please download the sample Common List document we created in this article:

*   [Common List.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Common%20List.pptx?raw=true)  
      
  

{{< alert style="warning" >}}Use this template for DB, Dataset, JSON and XML examples also.{{< /alert >}}

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 136c43e5f02469bb1cfbc494ed172ec9 >}}



#### Database Entities

{{< gist GroupDocsGists 1ca2bfad6924d932a1c861d0924db1d9 >}}



#### Using DataSet

{{< gist GroupDocsGists 413cbef26b96c3242d18fbdfc25f6d48 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 248a1f47e8bbb16f232a303b20e1c6b0 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists ee3e238e42381670c2672e4ef2556abd >}}



## Common List in OpenOffice Presentation Document

### Creating the Template

OpenDocument Presentation (ODP) is a presentation document format which can be used as an alternative to Microsoft PowerPoint Document (PPT/PPTX) formats. Since ODP is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODP, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating a 'Common List' report in ODP format. Instead, we'll save the existing template to ODP format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" dropdown.
4.  Click "Save".

### Download Template

*   [Common List\_OpenDocument.odp](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Common%20List_OpenDocument.odp?raw=true) (use same template for database, dataset, JSON and XML examples)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 44e2f9f2077892452ee6b31db0352150 >}}



#### Database Entities

{{< gist GroupDocsGists b1cc69c6c963bb480a2f847b7f93a63f >}}



#### Using DataSet

{{< gist GroupDocsGists a1e384eb5b528d6effe2a4699b4f211f >}}



#### Using XML DataSource

{{< gist GroupDocsGists 8164a98e56f424a01ba598ac2da41e0b >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 9298bdc957de3cccc19f5978f3bb9813 >}}



### ODP Template and Report in Apache OpenOffice

In order to check compatibility of ODP between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODP template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.

*   We opened the ODP report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
