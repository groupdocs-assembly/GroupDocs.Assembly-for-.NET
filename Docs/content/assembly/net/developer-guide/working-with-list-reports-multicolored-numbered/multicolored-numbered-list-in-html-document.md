---
id: multicolored-numbered-list-in-html-document
url: assembly/net/multicolored-numbered-list-in-html-document
title: Multicolored Numbered List in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Multicolored numbered list report in HTML Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

## Multicolored Numbered List in HTML Document

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in the numbered list.
*   It must highlight the products.
*   The report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:

<ol>

&lt;&lt;foreach [in products]&gt;&gt;
<li &lt;&lt;if [IndexOf() % 2 == 0]&gt;&gt;style="background-color:#FFF8DC"&lt;&lt;/if&gt;&gt;>&lt;&lt;[ProductName]&gt;&gt;</li>
&lt;&lt;/foreach&gt;&gt;
</ol>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Multicolored Numbered List Template

Please download the multicolored numbered list template document we created in this article:

*   [Multicolored Numbered List.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/Multicolored%20Numbered%20List.html?raw=true)

### Generating The Report

{{< gist GroupDocsGists 565c8c90f6d8cfe139534c6de81080cb >}}


