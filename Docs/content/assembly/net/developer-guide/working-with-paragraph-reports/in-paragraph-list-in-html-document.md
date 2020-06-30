---
id: in-paragraph-list-in-html-document
url: assembly/net/in-paragraph-list-in-html-document
title: In-Paragraph List in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Paragraph List report in HTML Document format.{{< /alert >}}

## In-Paragraph List in HTML Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   A descriptive or informative line like "We provide support for the following products:".
*   Show all the products along with the above sentence.
*   Report must be generated in the HTML Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products: <<foreach [in products]>>
<<[indexOf() != 0 ? ", " : ""]>><<[ProductName]>><</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download In-Paragraph List Template

Please download the sample Common List document we created in this article:

*   [In-Paragraph List.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/In-Paragraph%20List.html?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist rida-fatima-aspose ff54dc56129ac0115dd6f2d61dc9fa8c >}}


