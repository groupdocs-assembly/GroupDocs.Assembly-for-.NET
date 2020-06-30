---
id: barcode-image-generation-in-spreadsheet-document
url: assembly/net/barcode-image-generation-in-spreadsheet-document
title: Barcode Image Generation in Spreadsheet Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 3.2.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Barcode Generation in Microsoft Spreadsheet

### Template Syntax

*   Add a textbox to your template at the place where you want a barcode image to be inserted.
*   Add the following syntax:

```csharp
<<barcode [customer.Barcode] -qr>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Template

Get the template from here.

*   [Barcode.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Barcode.xlsx?raw=true)

### The Code

{{< gist GroupDocsGists fccb0d1f515216932a9c3844297d3d1e >}}


