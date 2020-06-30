---
id: bulleted-list-in-text-document
url: assembly/net/bulleted-list-in-text-document
title: Bulleted List in Text Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bulleted List report in Text Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Bulleted List in Text Document

{{< alert style="info" >}}This feature is supported by version 17.04 or greater{{< /alert >}}

### Creating a Bulleted List

Practicing the following steps you can insert Bulleted List in a Text document.

1.  To add a bulleted list use • as a bullet for each element in the list
2.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share the following key requirements:

*   A report must show all the products in a bulleted list format.
*   A report must be generated in the Text Document format.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>>•	<<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

### Download Bulleted List Template

Please download the sample Bulleted List document we created in this article:

*   [Bulleted List.txt](https://github.com/rida-fatima-aspose/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Bulleted%20List.txt?raw=true)

### Generating The Report

{{< gist GroupDocsGists f8f45cf046005387e1e121f3118cd9fb >}}


