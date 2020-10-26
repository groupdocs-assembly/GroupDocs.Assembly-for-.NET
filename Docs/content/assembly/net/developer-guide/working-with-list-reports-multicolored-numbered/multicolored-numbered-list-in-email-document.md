---
id: multicolored-numbered-list-in-email-document
url: assembly/net/multicolored-numbered-list-in-email-document
title: Multicolored Numbered List in Email Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Multicolored Numbered List report in Email Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

## Multicolored Numbered List in Email Document

### Creating a Multicolored Numbered List

Practising the following steps you can create Multicolored Numbered List Template in MS Outlook 2013.

1.  In your email, write a sentence like "We provide support for the following clients:".
2.  Start numbered list.
3.  Go to the "Design" tab and select color to make it colored list.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent the following key requirements:

*   The report must be in .eml or .msg format.
*   It must add email recipient, CSS and subject of the email.
*   The report must show the products in the numbered list.
*   It must highlight the products.
*   The report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

We provide support for the following products:

1. <<foreach \[in getClients()\]>><<if \[indexOf() % 2 == 0\]>>      <<\[numberOf()\]>>.    <<\[ProductName\]>>

2. &lt;&lt;else>>      &lt;&lt;[numberOf()]>>.      &lt;&lt;[ProductName]>>

<</if>><</foreach>>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

Download Multicolored Numbered List Template

Please download the sample Multicolored Numbered List document we created in this article:

*   [Multicolored Numbered List.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Multicolored%20Numbered%20List.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists b09c7249dba7302448473eb4f54ff510 >}}


