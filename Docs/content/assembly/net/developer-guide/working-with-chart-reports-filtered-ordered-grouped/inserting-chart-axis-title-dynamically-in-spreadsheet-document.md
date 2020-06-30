---
id: inserting-chart-axis-title-dynamically-in-spreadsheet-document
url: assembly/net/inserting-chart-axis-title-dynamically-in-spreadsheet-document
title: Inserting Chart Axis Title Dynamically in Spreadsheet Document
weight: 7
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 18.1 or greater.{{< /alert >}}

## Column Chart in Microsoft Excel Document

### Creating a Column Chart

Following steps, you can create a column chart in MS Excel 2013:

1.  Create a new Workbook
2.  Click the "Insert" tab, and then click "Insert Column Chart" icon in the Charts group to view the drop-down list
3.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your Worksheet
4.  Edit the Worksheet with your data to update the chart
5.  Save your Document

### Reporting Requirement

As a report developer, you are required to share orders quantity of the customers dynamically with the following key requirements:

*   A report must show the quantity of sales/orders
*   Sales/orders quantity must be represented by Quarters
*   It must associate order quantity with the corresponding customer
*   A report must be generated in the Spreadsheet Document

### Adding Syntax to be Evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
<<[title]>><<foreach [in orders
.Where(c => c.OrderDate.Year == 2015)
.GroupBy(c => c.Customer)
.OrderBy(g => g.Key.CustomerName)]>><<x [Key.CustomerName]>>

```

#### Chart Data

**Legend Entries**

```csharp
="1st Quarter<<y [Where(c => c.OrderDate.Month >= 1 && c.OrderDate.Month <= 3).Sum(c => c.ProductQuantity)]>>"
="2nd Quarter<<y [Where(c => c.OrderDate.Month >= 4 && c.OrderDate.Month <= 6).Sum(c => c.ProductQuantity)]>>"
="3rd Quarter<<y [Where(c => c.OrderDate.Month >= 7 && c.OrderDate.Month <= 9).Sum(c => c.ProductQuantity)]>>"
="4th Quarter<<y [Where(c => c.OrderDate.Month >= 10 && c.OrderDate.Month <= 12).Sum(c => c.ProductQuantity)]>>"

```

### Download Template

Please download the sample Chart with Filtering, Grouping, and Ordering with Dynamic Title document we created in this article:

*   [Chart Template.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_Dynamic_Title.xlsx) (Template for CustomObject and JSON examples) 

### Generating The Report

{{< gist GroupDocsGists 251fa6d077fd45b1222a6f352c4329b2 >}}


