---
id: numbered-list-in-html-document
url: assembly/net/numbered-list-in-html-document
title: Numbered List in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Numbered List report in HTML Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

## Numbered List in HTML Document

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in numbered list.
*   The report must be generated in the HTML Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<ol>
&lt;&lt;foreach [in products]&gt;&gt;
<li>&lt;&lt;[ProductName]&gt;&gt;</li>
&lt;&lt;/foreach&gt;&gt;
</ol>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Numbered List Template

Please download the sample Numbered List document we created in this article:

*   [Numbered List.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/Numbered%20List.html?raw=true)

### Generating The Report

{{< gist GroupDocsGists b36160effc7441adc59a543ce5934fff >}}


