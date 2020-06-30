---
id: in-table-master-detail-in-html-document
url: assembly/net/in-table-master-detail-in-html-document
title: In-Table Master-Detail in HTML Document
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate In-Table Master-Detail report in HTML Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.03 or greater.{{< /alert >}}

## In-Table Master-Detail in HTML Document

### Reporting Requirement

As a report developer, you are required to represent customers' total orders prices and price of each product separately with the following key requirements:

*   The report must show each customer along with his total orders prices.
*   It must also show each individual product ordered by the customers.
*   It must show sum of the order prices.
*   It must represent all the information in tabular form.
*   The report must be generated in the HTML Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<<foreach \[in customers\]>> <<foreach \[in Order\]>> <</foreach>> <</foreach>>

| Customer | Order Price |
| --- | --- |
| **<<\[CustomerName\]>>** | <<\[Order.Sum(c => c.Price)\]>> |
| *  <<\[Product.ProductName\]>>* | <<\[Price\]>> |
| **Total:** | <<\[Sum( m => m.Order.Sum( c => c.Price))\]>> |

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine.{{< /alert >}}

Download  In-Table Master-Detail Template

Please download the sample In-Table Master-Detail document we created in this article:

*   [In-Table Master-Detail.html](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/HTML%20Templates/In-Table%20Master-Detail.html?raw=true)

### Generating The Report

{{< gist GroupDocsGists 57a8db2df722098fc33343718b289296 >}}


