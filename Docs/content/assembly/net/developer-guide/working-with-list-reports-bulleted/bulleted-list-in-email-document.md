---
id: bulleted-list-in-email-document
url: assembly/net/bulleted-list-in-email-document
title: Bulleted List in Email Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bulleted List report in Email format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Bulleted List in Email

{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

#### Reporting Requirement

As a report developer, you are required to share the following key requirements:

*   A report must be in .eml or .msg format.
*   It must add email recipient, CSS and subject of the email.
*   A report must show all the products in a bulleted list format.

#### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

##### Recipients, Cc and Subject

```csharp
<<foreach [r in recipients]>><<[r]>>; <</foreach>>
<<[cc]>>
<<[subject]>>

```

##### Body

```csharp
We provide support for the following clients:
<<foreach[in products]>>  ∙  <<[ProductName]>>
<</foreach>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

#### Download Bulleted List Template

Please download the sample Bulleted List document we used in this article:

*   [Bulleted List.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Bulleted%20List.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists fd6c17536da63fea2bb7f8923a798993 >}}


