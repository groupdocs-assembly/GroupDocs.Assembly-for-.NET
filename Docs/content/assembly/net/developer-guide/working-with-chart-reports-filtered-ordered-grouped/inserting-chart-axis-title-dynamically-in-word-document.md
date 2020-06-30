---
id: inserting-chart-axis-title-dynamically-in-word-document
url: assembly/net/inserting-chart-axis-title-dynamically-in-word-document
title: Inserting Chart Axis Title Dynamically in Word Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.12 or greater.{{< /alert >}}

## Column Chart in Microsoft Word Document

### Creating a Column Chart

Practising the following steps, you can insert a Column Chart in MS Word 2013:

1.  Click in the document where you want to insert the chart, click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box
2.  Select "Column" in the sidebar, you will see a gallery of charts
3.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your document
4.  Edit the Worksheet with your data to update the chart. See [Chart Data (Excel)]({{< ref "assembly/net/developer-guide/working-with-chart-reports-filtered-ordered-grouped/inserting-chart-axis-title-dynamically-in-word-document.md" >}})
5.  Save your Document

### Reporting Requirement

As a report developer, you are required to share orders quantity of the customers dynamically with the following key requirements:

*   A report must show the quantity of sales/orders
*   Sales/orders quantity must be represented by Quarters
*   It must associate order quantity with the corresponding customer
*   A report must be generated in the Word Processing Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
<<title>><<foreach [in orders
.Where(c => c.OrderDate.Year == 2015)
.GroupBy(c => c.Customer)
.OrderBy(g => g.Key.CustomerName)]>><<x [Key.CustomerName]>>

```

#### Chart Data (Excel)

|   | 1st Quarter<<y[Where(c => c.OrderDate.Month>= 1 && c.OrderDate.Month <= 3).Sum(c => c.ProductQuantity)]>> | 2nd Quarter<<y[Where(c => c.OrderDate.Month>= 4 && c.OrderDate.Month <= 6).Sum(c => c.ProductQuantity)]>> | 3rd Quarter<<y[Where(c => c.OrderDate.Month>= 7 && c.OrderDate.Month <= 9).Sum(c => c.ProductQuantity)]>> | 4th Quarter<<y[Where(c => c.OrderDate.Month>= 10 && c.OrderDate.Month <= 12).Sum(c => c.ProductQuantity)]>> |
| --- | --- | --- | --- | --- |
| Category 1 | 4.3 | 2.4 | 2 | 3 |
| Category 2 | 2.5 | 4.4 | 2 | 2 |
| Category 3 | 3.5 | 1.8 | 3 | 5 |
| Category 4 | 4.5 | 2.8 | 5 | 2 |

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Template

Please download the sample Chart with Filtering, Grouping, and Ordering document we created in this article:

*   [Chart Template.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_dynamic_title.docx) (Template for CustomObject and JSON examples) 

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists df2cabaaed2bd0016d8e86ed74647366 >}}


