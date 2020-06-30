---
id: single-row-image-in-email-document
url: assembly/net/single-row-image-in-email-document
title: Single Row Image in Email Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Single Row Image report in Email Document format.{{< /alert >}}

## Single Row Image in Email Document

{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent information of first single manager with the following key requirements:

*   Report must be in .eml or .msg format
*   It must add email recipient, css and subject of the email
*   Report must show image of the manager
*   It must show Name and contact number of customer

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
Name:	<<[customer.CustomerName]>>
Contact Number:	<<[customer.CustomerContactNumber]>>

```

### Download Single Row Image Template

Please download the sample template we created in this article:

*   [Single Row Image.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Single%20Row.msg?raw=true)

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist rida-fatima-aspose d1b17449056fb2f542e302e6015014c8 >}}


