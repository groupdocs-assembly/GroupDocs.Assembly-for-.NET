---
id: bubble-chart-in-presentation-document
url: assembly/net/bubble-chart-in-presentation-document
title: Bubble Chart in Presentation Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Bubble Chart in Microsoft PowerPoint Document

### Creating a Bubble Chart

Please follow below steps to create Bubble Chart in MS PowerPoint 2013:

1.  Add a new presentation slide, click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box.
2.  Select "X Y(Scatter)" in the sidebar, you will see a gallery of charts.
3.  Select the "Bubble" and press "OK" to insert the chart and Worksheet template to your document.
4.  Edit the Worksheet with your data to update the chart. See [Chart Data (Excel)](https://docs.groupdocs.com/assembly/net/bubble-chart-in-presentation-document/#adding-syntax-to-be-evaluated-by-groupdocsassembly-engine).
5.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share your sales/orders dynamically with the following key requirements:

*   A report must show variation in the sales/orders prices.
*   It must show which sale/order lie in which month.
*   A report must be generated in the Presentation Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Orders Prices by Months<<foreach [in orders
.GroupBy(c => c.OrderDate.Month)]>><<x [Key]>>

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

#### Chart Data (Excel)

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 47.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">X-Values</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 328.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">Total orders by months&lt;&lt;y [Sum(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;&lt;&lt;size [Count()]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 37.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">Size</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 47.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">0.7</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 328.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">2.7</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 37.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">10</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 47.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">1.8</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 328.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">3.2</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 37.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">4</span></p></td></tr><tr style="height: 15pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 47.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">2.6</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 328.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">0.8</span></p></td><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 37.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: right;"><span style="font-family: Calibri; font-size: 11pt;">8</span></p></td></tr></tbody></table>
### Download Bubble Chart Template

Please download the sample Bubble Chart document we created in this article:

*   [Bubble Chart.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Bubble%20Chart.pptx?raw=true) (Template for CustomObject and JSON examples) 
*   [Bubble Chart\_DB.pptx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Presentation%20Templates/Bubble%20Chart_DB.pptx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists c2e605d2d5fba00340f51a2e09a5bb9e >}}



#### Database Entities

{{< gist GroupDocsGists f1e9da2dfb1bbafc6437762412dc6653 >}}



#### Using DataSet

{{< gist GroupDocsGists 64eab241b98c0edad88f177e92513d8b >}}



#### Using XML DataSource

{{< gist GroupDocsGists 2a372988c264be1acfa5fd8848fc9fc6 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 7be60df67524b051e8dda43435d0b69e >}}



## Bubble Chart in OpenOffice Presentation Document

To be investigated.
