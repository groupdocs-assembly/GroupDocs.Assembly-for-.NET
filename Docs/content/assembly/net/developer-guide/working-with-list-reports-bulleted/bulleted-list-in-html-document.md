---
id: bulleted-list-in-html-document
url: assembly/net/bulleted-list-in-html-document
title: Bulleted List in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bulleted List report in HTML Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Bulleted List in HTML Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

### Creating a Bulleted List

Practising the following steps you can insert Bulleted List in an HTML document.

1.  To add a bulleted list start and end your list with `<ul></ul>` tag and list all the elements by enclosing them between `<li></li>` tag.
2.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share the following key requirements:

*   A report must show all the products in a bulleted list format.
*   A report must be generated in the HTML Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<ul>
&lt;&lt;foreach [in products]&gt;&gt;
<li>&lt;&lt;[ProductName]&gt;&gt;</li>
&lt;&lt;/foreach&gt;&gt;
</ul>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Bulleted List Template

Please download the sample Bulleted List document we created in this article:

*   [Bulleted List.html](https://github.com/rida-fatima-aspose/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/Bulleted%20List.html?raw=true)

### Generating The Report

{{< gist GroupDocsGists fb223b7e519191219fec1e9499c8e9d8 >}}


