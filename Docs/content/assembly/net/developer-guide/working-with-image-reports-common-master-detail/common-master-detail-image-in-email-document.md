---
id: common-master-detail-image-in-email-document
url: assembly/net/common-master-detail-image-in-email-document
title: Common Master-Detail Image in Email Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail report in Email Document format.{{< /alert >}}

## Common Master-Detail in Email Document

{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater{{< /alert >}}

### Creating a Common Master-Detail

Please follow below steps to create Common Master-Detail Template in MS Outlook 2013:

1.  Create a new Email.
2.  Insert two shapes, one for holding image and other for holding other details.
3.  Save the Email.

### Reporting Requirement

As a report developer, you are required to represent the information of the managers and clients with the following key requirements:

*   Report must be in .eml or .msg format.
*   It must add email recipient, css and subject of the email.
*   Report must show customer's name.
*   It must associate the customers with their products.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
<<foreach [in customers]>>Customer:
<<[CustomerName]>>
Products: <<foreach [in
Order]>><<[IndexOf() != 0 ? ", " : ""]>><<[Product.ProductName]>><</foreach>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Common Master-Detail Template

Please download the sample Common Master-Detail document we created in this article:

*   [Common Master-Detail.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Common%20Master-Detail.msg?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist rida-fatima-aspose 25a382242465c5c3305a33d90e7aa1b4 >}}


