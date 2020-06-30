---
id: pie-chart-in-spreadsheet-document
url: assembly/net/pie-chart-in-spreadsheet-document
title: Pie Chart in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Pie Chart report in Spreadsheet Document format based on the use case: Working with a Business Case.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Pie Chart in Microsoft Excel Document

### Creating a Pie Chart

Please follow below steps to create Pie Chart in MS Excel 2013.

1.  Add a new Workbook.
2.  Click in the workbook where you want to insert the chart, click the "Insert" tab, and then click "Pie Chart Icon" in the charts group.
3.  A drop-down with charts will appear, select the "Pie" and press "OK" to insert the chart.
4.  Click on the chart you just inserted, then click the "Change Data" icon in Data group.
5.  Now add legend entries. See [Chart Data]({{< ref "assembly/net/developer-guide/working-with-chart-reports-pie/pie-chart-in-spreadsheet-document.md" >}})
6.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share your customers' orders details dynamically with the following key requirements:

*   The report must show information on a Pie Chart.
*   It must indicate customer name with value (price of the orders purchased).
*   The report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Price<<foreach [in customers]>>
<<x [CustomerName]>>

```

#### Chart Data

**Legend Entries**

```csharp
="Total Order Price<<y [Order.Sum(c => c.Price)]>>"

```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

Download Pie Chart Template

Please download the sample Pie Chart document we created in this article:

*   [Pie Chart.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Pie%20Chart.xlsx?raw=true) (Template for CustomObject and JSON examples) 
*   [Pie Chart\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Pie%20Chart_DB.xlsx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 9fbffedfe245612403cffebea41c2191 >}}



#### Database Entities

{{< gist GroupDocsGists 56046573d12769a1b23fed18850cf8f9 >}}



#### Using DataSet

{{< gist GroupDocsGists 885ade4670358b72fcfb00fa9bf5fe5e >}}



#### Using XML DataSource

{{< gist GroupDocsGists 0cb24748eee81d88b205c0f19b2b4f62 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 1372b2b5fb02b3ac72ceb69a023d3b17 >}}



## Pie Chart in OpenOffice Spreadsheet Document

To be investigated.
