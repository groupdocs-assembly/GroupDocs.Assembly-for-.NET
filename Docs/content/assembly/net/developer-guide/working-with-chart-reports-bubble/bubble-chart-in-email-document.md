---
id: bubble-chart-in-email-document
url: assembly/net/bubble-chart-in-email-document
title: Bubble Chart in Email Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bubble Chart report in Email Document.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Bubble Chart in Email Document

{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

### Reporting Requirement

As a report developer, you are required to share your sales/orders dynamically with the following key requirements:

*   A report must be in .eml or .msg format.
*   It must add email recipient, CSS and subject of the email.
*   Retrieve total order Prices by Months.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

##### Chart Title

```csharp
Orders Prices by Months<<foreach [in orders
.GroupBy(c => c.OrderDate.Month)]>><<x[Key]>>

```

##### Chart Data (Excel)

<table border="1" cellspacing="0" cellpadding="0" width="623" style="width: 467.5pt; border-collapse: collapse; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"><tbody><tr><td width="108" valign="top" style="width: 80.75pt; border-top-color: windowtext; border-top-style: solid; border-top-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: windowtext; border-left-style: solid; border-left-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p style="margin-bottom: 0.0001pt; line-height: normal;"><span style="color: black;">X-Values</span></p></td><td width="438" valign="top" style="width: 328.5pt; border-top-color: windowtext; border-top-style: solid; border-top-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: initial; border-left-style: none; border-left-width: initial; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p style="margin-bottom: 0.0001pt; line-height: normal;"><span style="color: black;">Total Order Price&lt;&lt;y [sum(c =&gt; c.Price)]&gt;&gt;&lt;&lt;size[Count()]&gt;&gt;</span></p></td><td width="78" valign="top" style="width: 58.25pt; border-top-color: windowtext; border-top-style: solid; border-top-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: initial; border-left-style: none; border-left-width: initial; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p style="margin-bottom: 0.0001pt; line-height: normal;"><span style="color: black;">Size</span></p></td></tr><tr><td width="108" valign="top" style="width: 80.75pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: windowtext; border-left-style: solid; border-left-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">0.7</span></p></td><td width="438" valign="top" style="width: 328.5pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">2.7</span></p></td><td width="78" valign="top" style="width: 58.25pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">10</span></p></td></tr><tr><td width="108" valign="top" style="width: 80.75pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: windowtext; border-left-style: solid; border-left-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">1.8</span></p></td><td width="438" valign="top" style="width: 328.5pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">3.2</span></p></td><td width="78" valign="top" style="width: 58.25pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">4</span></p></td></tr><tr><td width="108" valign="top" style="width: 80.75pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: windowtext; border-left-style: solid; border-left-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">2.6</span></p></td><td width="438" valign="top" style="width: 328.5pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">0.8</span></p></td><td width="78" valign="top" style="width: 58.25pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial; border-bottom-color: windowtext; border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: windowtext; border-right-style: solid; border-right-width: 1pt; padding-top: 0in; padding-right: 5.4pt; padding-bottom: 0in; padding-left: 5.4pt;"><p align="right" style="margin-bottom: 0.0001pt; text-align: right; line-height: normal;"><span style="color: black;">8</span></p></td></tr></tbody></table>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

### Download Bubble Chart Template

Please download the sample Bubble Chart document we used in this article:

*   [Bubble Chart.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Bubble%20Chart.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists 9fce9ede8837b3240df371a0eb596a67 >}}


