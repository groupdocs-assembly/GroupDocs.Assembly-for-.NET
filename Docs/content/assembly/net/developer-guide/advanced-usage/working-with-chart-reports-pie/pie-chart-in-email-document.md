---
id: pie-chart-in-email-document
url: assembly/net/pie-chart-in-email-document
title: Pie Chart in Email Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Pie Chart report in Presentation Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

## Pie Chart in Email Document

### Creating a Pie Chart

Please follow below steps to create Pie Chart in MS Outlook 2013:

1.  Create a new Email.
2.  Click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box.
3.  Select "Pie".
4.  Preview "Pie" and press OK to insert the chart and Worksheet template to your email.
5.  Edit the Worksheet with your data to update the chart.

### Reporting Requirement

As a report developer, you are required to share managers' contract prices dynamically with the following key requirements:

*   The report must be in .eml or .msg format.
*   It must add email recipient, CSS and subject of the email.
*   The report must show information on a Pie Chart.
*   It must indicate customer name with value (price of the order).

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

##### Chart Title

```csharp
Total Order Price<<foreach [in
customers]>><<x [CustomerName]>>
```

##### Chart Data (Excel)

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 219.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">Total Order Price&lt;&lt;y [</span><span style="font-family: Calibri; font-size: 11pt;">Order.Sum</span><span style="font-family: Calibri; font-size: 11pt;">(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 219.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">8.2</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 219.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">3.2</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 219.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">1.4</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 219.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">1.2</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Pie Chart Template

Please download the sample Pie Chart document we used in this article:

*   [Pie Chart.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Pie%20Chart.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists 16f8e9bddf60f4a95459a2898d9357aa >}}


