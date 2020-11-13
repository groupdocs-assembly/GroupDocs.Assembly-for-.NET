---
id: common-master-detail-in-text-document
url: assembly/net/common-master-detail-in-text-document
title: Common Master-Detail in Text Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail report in Text Document format.{{< /alert >}}

## Common Master-Detail in Text Document

{{< alert style="info" >}}This feature is supported by version 17.3.0 or greater.{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   It must associate the customers with their products
*   Report must be generated in the HTML Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
<<foreach [in customers]>><<[CustomerName()]>>
Products: <<foreach [in Order]>><<[indexOf() != 0 ? ", " : ""]>><<[Product.ProductName]>><</foreach>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Common Master-Detail Template

Please download the sample Common master-detail document we created in this article:

*   [Common Master-Detail.txt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Common%20Master-Detail.txt?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist rida-fatima-aspose 5101b2428d3e7fd66bf4ba71d28c3967 >}}


