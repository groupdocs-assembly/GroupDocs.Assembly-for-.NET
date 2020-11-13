---
id: column-chart-in-spreadsheet-document
url: assembly/net/column-chart-in-spreadsheet-document
title: Column Chart in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Column Chart Report with Filtered, Ordered and Grouped Data in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Column Chart in Microsoft Excel Document

### Creating a Column Chart

Following below steps, you can create a column chart in MS Excel 2013:

1.  Create a new Workbook.
2.  Click the "Insert" tab, and then click "Insert Column Chart" icon in the Charts group to view the drop-down list.
3.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your Worksheet.
4.  Edit the Worksheet with your data to update the chart. See [Chart Data](https://docs.groupdocs.com/assembly/net/column-chart-in-spreadsheet-document/#adding-syntax-to-be-evaluated-by-groupdocsassembly-engine).
5.  Save your Document.

### Reporting Requirement

As a report developer, you are required to share orders quantity of the customers dynamically with the following key requirements:

*   A report must show the quantity of sales/orders.
*   Sales/orders quantity must represented by Quarters.
*   It must associate order quantity with the corresponding customer.
*   A report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Quantity by Quarters<<foreach [in orders
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

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Template

Please download the sample Chart with Filtering, Grouping, and Ordering document we created in this article:

*   [Chart Template.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering.xlsx?raw=true) (Template for CustomObject and JSON examples) 
*   [Chart Template\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_DB.xlsx?raw=true) (Template for DataSet, DataBase examples)
*   [Chart Template\_XML.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_XML.xlsx?raw=true) (Template for XML examples) 

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 2617acf16e4c9ab06ebc84b5be9c16b3 >}}



#### Database Entities

{{< gist GroupDocsGists a7b411f1f03bb75b071a2603f153fade >}}



#### Using DataSet

{{< gist GroupDocsGists 188f8d564c951ce7aafa24077e429d75 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 12e4c899cf368461f801f0b3fe936496 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 03f52003d74ce0002c880de8c90fcd91 >}}



## Column Chart in OpenOffice Spreadsheet Document

To be investigated.
