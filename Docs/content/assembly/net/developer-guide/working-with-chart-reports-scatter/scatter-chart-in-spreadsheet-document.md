---
id: scatter-chart-in-spreadsheet-document
url: assembly/net/scatter-chart-in-spreadsheet-document
title: Scatter Chart in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Scatter Chart report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Scatter Chart in Microsoft Excel Document

### Creating a Scatter Chart

Please follow below steps to create Scatter Chart in MS Excel 2013:

1.  Add a new Workbook.
2.  Click in the workbook where you want to insert the chart, click the "Insert" tab, and then click "Insert Scatter Chart Icon" in the charts group.
3.  A drop-down with charts will appear, select the "Scatter" and press "OK" to insert the chart.
4.  Click on the chart you just inserted, then click the "Change Data" icon in Data group.
5.  Now add legend entries. See [Chart Data]({{< ref "assembly/net/developer-guide/working-with-chart-reports-scatter/scatter-chart-in-spreadsheet-document.md" >}}).
6.  Save your Document.

### Reporting Requirement

As a report developer, you are required to show your customers' orders prices by month with the following key requirements:

*   Report must show information on a Scatter Chart.
*   It must indicate Price of each order by month.
*   Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Prices by Months<<foreach [in orders
.GroupBy(c => c.OrderDate.Month)]>>
```

#### Chart Data

**Legend Entries**

```csharp
="Total Order Price<<x [Key]>><<y [Sum(c => c.Price)]>>"
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Scatter Chart Template

Please download the sample Scatter Chart document we created in this article:

*   [Scatter Chart.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Scatter%20Chart.xlsx?raw=true) (Template for CustomObject and JSON examples) 
*   [Scatter Chart\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Scatter%20Chart_DB.xlsx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 2dfb4c8b14814612e05d5d144887df2e >}}



#### Database Entities

{{< gist GroupDocsGists 9c0bb7a58a06fc2eaaa780af1e51ee13 >}}



#### Using DataSet

{{< gist GroupDocsGists 4c2770bf096eb8b3670433b695610e5f >}}



#### Using XML DataSource

{{< gist GroupDocsGists 55bb62b4b8afd5c31ac08452352d8197 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 3ec160ae6172521393ec72139f8af94c >}}



## Scatter Chart in OpenOffice Spreadsheet Document

To be investigated.
