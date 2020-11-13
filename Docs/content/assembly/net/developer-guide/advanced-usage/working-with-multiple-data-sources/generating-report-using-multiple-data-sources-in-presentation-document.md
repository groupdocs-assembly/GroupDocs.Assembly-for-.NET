---
id: generating-report-using-multiple-data-sources-in-presentation-document
url: assembly/net/generating-report-using-multiple-data-sources-in-presentation-document
title: Generating Report using Multiple Data Sources in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a report. This report will fetch data from multiple data sources.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Generating Report in Microsoft PowerPoint Document

### Creating a Template

1.  Add a new slide
2.  Add a bullet list at the place where you want it
3.  Save your Document

## Reporting Requirement

As a report developer, you are required to generate a report that fetches data from two different data sources (e:g XML, JSON). Report must show following information:

*   Customer name
*   List of products

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

{{< alert style="warning" >}}There is no way to use an image inside a foreach tag in Microsoft PowerPoint.{{< /alert >}}

**NOTE: There is no way to use images inside a foreach tag in PowerPoint.**

<<foreach \[in customers\]>>**<<\[CustomerName\]>>**

<</foreach\>>

```csharp
We provide support for the following products:
<<foreach [in products]>><<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download Template

Get the template from here.

*   [Multiple DS.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Multiple%20DS.pptx?raw=true)

### Generating the Report

{{< gist GroupDocsGists d9bf0942465e757e9db871545d904459 >}}

## Generating Report in OpenOffice Presentation Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating a report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

Get template from here.

*   [Multiple DS.odp](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Multiple%20DS.odp?raw=true)

### Generating the Report

{{< gist GroupDocsGists 296b38df6d5eb9a8c31fded8617b9739 >}}


