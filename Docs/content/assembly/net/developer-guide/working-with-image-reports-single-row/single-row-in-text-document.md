---
id: single-row-in-text-document
url: assembly/net/single-row-in-text-document
title: Single Row in Text Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Single Row report in Text Document format.{{< /alert >}}

## Single Row in Text Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent information of first single customer with the following key requirements:

*   It must show Name and Contact Number of the customer
*   Report must be generated in the Text Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
Name:	<<[customer.CustomerName]>>
Contact Number:	<<[customer.CustomerContactNumber]>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Single Row Image Template

Please download the single row image document we created in this article:

*   [Single Row.txt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Single%20Row.txt?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< gist GroupDocsGists 2e51cfff8166768ebdcb9fa4de93e251 >}}


