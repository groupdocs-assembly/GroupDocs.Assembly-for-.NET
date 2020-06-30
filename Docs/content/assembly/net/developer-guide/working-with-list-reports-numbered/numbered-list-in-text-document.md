---
id: numbered-list-in-text-document
url: assembly/net/numbered-list-in-text-document
title: Numbered List in Text Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Numbered List report in Text Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

## Numbered List in Text Document

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in the numbered list.
*   The report must be generated in the Text Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[NumberOf()]>>. <<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

Download Numbered List Template

Please download the sample Numbered List document we created in this article:

*   [Numbered List.txt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Numbered%20List.txt?raw=true)

### Generating The Report

{{< gist GroupDocsGists df9028fecf23b9a87f197c8cfc8c2208 >}}


