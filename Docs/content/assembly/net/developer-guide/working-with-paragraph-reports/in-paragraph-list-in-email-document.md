---
id: in-paragraph-list-in-email-document
url: assembly/net/in-paragraph-list-in-email-document
title: In-Paragraph List in Email Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Paragraph List report in Email Document format.{{< /alert >}}

## In-Paragraph List in Email Document

{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to share the following key requirements:

*   Report must be in .eml or .msg format.
*   It must add email recipient, css and subject of the email.
*   A descriptive or informative line like "We provide support for the following products:".
*   Show all the products along with the above sentence.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[IndexOf() != 0 ? ", " : ""]>>
<<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download In-Paragraph List Template

Please download the sample Common List document we created in this article:

*   [In-Paragraph List.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/In-Paragraph%20List.msg?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist rida-fatima-aspose b447f473d018009724f467a305acbc52 >}}


