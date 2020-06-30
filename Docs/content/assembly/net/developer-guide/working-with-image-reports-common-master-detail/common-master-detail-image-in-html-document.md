---
id: common-master-detail-image-in-html-document
url: assembly/net/common-master-detail-image-in-html-document
title: Common Master-Detail Image in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail report in HTML Document format.{{< /alert >}}

## Common Master-Detail Image in HTML Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   Report must show customers' picture and name.
*   It must associate the customers with their products.
*   Report must be generated in the HTML Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<<foreach \[in customers\]>>

![](data:image/jpeg;base64,<<[Photo]>>)

**<<\[CustomerName\]>>**

**Products:**

**

<<foreach \[in Order\]>><<\[IndexOf() != 0 ? ", " : ""\]>>

<<\[Product.ProductName\]>>

**

**<</foreach>>**

<</foreach>>

 {{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Common Master-Detail Template

Please download the sample Common master-detail document we created in this article:

*   [Common Master-Detail.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/Common%20Master-Detail.html?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist rida-fatima-aspose 31319a8a70569c70ad671106eeb1404b >}}


