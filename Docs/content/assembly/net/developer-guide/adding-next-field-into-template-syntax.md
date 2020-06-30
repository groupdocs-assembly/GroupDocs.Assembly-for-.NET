---
id: adding-next-field-into-template-syntax
url: assembly/net/adding-next-field-into-template-syntax
title: Adding NEXT Field into Template Syntax
weight: 32
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 3.3.0 or later releases.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Adding NEXT Field into Template Syntax

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

**Template Syntax**

```csharp
We provide support for the following products:
. <<foreach [in products]>><<next>><<[ProductName]>>
<</foreach>>

```

### The Recipe

*   Set up the source document template path
*   Set up destination report path
*   Instantiate [DocumentAssembler](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassembler) class
*   Generate report

### Download Template

Get the template from here.

*   [Using Next.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Next.docx?raw=true)

### The Code

{{< gist GroupDocsGists 5b2f5abf4f1a745b99baba52177f4db3 >}}


