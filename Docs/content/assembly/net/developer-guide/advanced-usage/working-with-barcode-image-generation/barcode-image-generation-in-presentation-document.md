---
id: barcode-image-generation-in-presentation-document
url: assembly/net/barcode-image-generation-in-presentation-document
title: Barcode Image Generation in Presentation Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 3.2.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Barcode Generation in Microsoft PowerPoint

### Template Syntax

*   Add a textbox to your template at the place where you want a barcode image to be inserted.
*   Add the following syntax:

```csharp
<<barcode [customer.Barcode] -codabar>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download Template

Get the template from here.

*   [Barcode.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Barcode.pptx?raw=true)

### The Code

{{< gist GroupDocsGists d11570805209774badf3def8336ae84a >}}
