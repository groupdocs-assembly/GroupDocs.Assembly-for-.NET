---
id: barcode-image-generation-in-word-processing-document
url: assembly/net/barcode-image-generation-in-word-processing-document
title: Barcode Image Generation in Word Processing Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 3.2.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Barcode Generation in Microsoft Word

### Template Syntax

*   Add a textbox to your template at the place where you want a barcode image to be inserted.
*   Add the following syntax:

```csharp
<<barcode [customer.Barcode] -codabar>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Template

Get the template from here.

*   [Barcode.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Barcode.docx?raw=true)

### The Code

{{< gist GroupDocsGists 0d42e27e26b1798117e9885319c8676b >}}


