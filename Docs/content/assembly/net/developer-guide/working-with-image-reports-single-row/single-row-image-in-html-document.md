---
id: single-row-image-in-html-document
url: assembly/net/single-row-image-in-html-document
title: Single Row Image in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Single Row Image report in HTML Document format.{{< /alert >}}

## Single Row Image in HTML Document

{{< alert style="info" >}}This feature is supported by version 17.03 or greater{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to represent information of first single customer with the following key requirements:

*   Report must show image of the customer
*   It must show Name and Contact Number of the customer
*   Report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
<table>
<tr>
<td>
<img height="100" width="100" src="data:image/jpeg;base64,&lt;&lt;[customer.Photo]&gt;&gt;"/>
</td>
<td>
<table>
<tr>
<td>
<b>Name:</b>
</td>
<td>
<b>&lt;&lt;[customer.CustomerName]&gt;&gt;</b>  
</td>
</tr>
<tr>
<td>
<b>Contact Number:</b>
</td>
<td>
&lt;&lt;[customer.CustomerContactNumber]&gt;&gt;
</td>
</tr>
</table>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Single Row Image Template

Please download the single row image document we created in this article:

*   [Single Row.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Text%20Templates/Numbered%20List.txt?raw=true)

### Generating The Report

{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer{{< /alert >}}{{< gist GroupDocsGists 9bbee318cdfc682815c7284ada1d94bf >}}


