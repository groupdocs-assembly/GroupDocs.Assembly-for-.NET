---
id: common-list-image-in-email-document
url: assembly/net/common-list-image-in-email-document
title: Common List Image in Email Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common List report in Email Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

## Common List in Email Document

### Reporting Requirement

As a report developer, you are required to share the following key requirements:

*   A report must be in .eml or .msg format
*   It must add email recipient, CSS and subject of the email
*   A report must show multiple customers' name

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
<<foreach [in customers]>><<[CustomerName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

Download Common List Template

Please download the sample Common List document we used in this article:

*   [Common List.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Common%20List.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists 1980c861285d1fd9be006efbc4a4a5c0 >}}


