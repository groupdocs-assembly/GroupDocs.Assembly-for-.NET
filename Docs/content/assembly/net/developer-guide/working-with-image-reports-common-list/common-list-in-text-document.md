---
id: common-list-in-text-document
url: assembly/net/common-list-in-text-document
title: Common List in Text Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common List report in Text Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

## Common List in Text Document

### Reporting Requirement

As a report developer, you are required to share the following key requirements:

*   The report must be generated in the Text Document format

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
<<foreach [in customers]>><<[CustomerName]>>
<</foreach>>
```

### Download Common List Template

Please download the sample Common List document we created in this article:

*   [Common List.txt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Common%20List.txt?raw=true)

### Generating The Report

{{< gist GroupDocsGists bd8aa3bbc25b483bde5e05c376ed8509 >}}


