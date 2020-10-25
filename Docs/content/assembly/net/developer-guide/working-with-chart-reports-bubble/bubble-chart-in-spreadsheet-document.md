---
id: bubble-chart-in-spreadsheet-document
url: assembly/net/bubble-chart-in-spreadsheet-document
title: Bubble Chart in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Bubble Chart in Microsoft Excel Document

### Creating a Bubble Chart

Please follow below steps to create a Bubble Chart in MS Excel 2013:

1.  Add a new Workbook.
2.  Click on the workbook where you want to insert the chart, click the "Insert" tab, and then click "Insert Scatter Chart Icon" in the charts group.
3.  A drop-down with charts will appear, select the "Bubble" and press "OK" to insert the chart.
4.  Click on the chart you just inserted, then click the "Select Data" icon in Data group.
5.  Add legend entries. See [Chart Data]({{< ref "assembly/net/developer-guide/working-with-chart-reports-bubble/bubble-chart-in-spreadsheet-document.md" >}}).
6.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share your sales/orders dynamically with the following key requirements:

*   A report must show variation in the sales/orders prices.
*   It must show which sale/order lie in which month.
*   A report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Orders Prices by Months<<foreach [in orders
.GroupBy(c => c.OrderDate.Month)]>><<x [Key]>>

```

#### Chart Data

**Legend Entries**

```csharp
="Orders Prices by Months<<y [Sum(c=> c.Price)]>><<size [Count()]>>"

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Bubble Chart Template

Please download the sample Bubble Chart document we created in this article:

*   [Bubble Chart.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Bubble%20Chart.xlsx?raw=true) (Template for CustomObject and JSON examples) 
*   [Bubble Chart\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Bubble%20Chart_DB.xlsx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 765476d18fe7954c79095567ef3a75ee >}}



#### Database Entities

{{< gist GroupDocsGists 2ac7c9d251e3d1a616bb937df12ccf7e >}}



#### Using DataSet

{{< gist GroupDocsGists bb0a55e6e06a556549797c43257f5ca0 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 4de723c02ab1c71ddda43bb5e6f6b22a >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 473a83fcdc1ecb0e5357628008361e09 >}}



## Bubble Chart in OpenOffice Spreadsheet Document

To be investigated.
