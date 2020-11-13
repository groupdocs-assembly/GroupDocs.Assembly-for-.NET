---
id: inserting-chart-axis-title-dynamically-in-email-document
url: assembly/net/inserting-chart-axis-title-dynamically-in-email-document
title: Inserting Chart Axis Title Dynamically in Email Document
weight: 9
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 18.1 or greater.{{< /alert >}}

## Column Chart in Email Document

### Creating a Column Chart

Please follow below steps to create a column chart in MS Outlook 2013:

1.  Create a new Email
2.  Click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialogue box
3.  Select "Column" in the sidebar, you will see a gallery of charts
4.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your email
5.  Edit the Worksheet with your data to update the chart
6.  Save your Email

### Reporting Requirement

As a report developer, you are required to share orders quantity of the customers dynamically with the following key requirements:

*   A report must be in .eml or .msg format
*   It must add email recipient, CSS and subject of the email
*   Report must show total Order Quantity by Quarters

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
<<[title>><<foreach
[in orders
.Where(c => c.OrderDate.Year == 2015)
.GroupBy(c => c.Customer)
.OrderBy(g => g.Key.CustomerName)]>><<x
[Key.CustomerName]>>
```

#### Chart Data (Excel)

|            | 1st Quarter<<y [Where(c => c.Date.Month >= 1 && c.Date.Month <= 3).Sum(c => c.ProductQuantity)]>> | 2nd Quarter<<y [Where(c => c.Date.Month >= 4 && c.Date.Month <= 6).Sum(c => c.ProductQuantity)]>> | 3rd Quarter<<y [Where(c => c.Date.Month >= 7 && c.Date.Month <= 9).Sum(c => c.ProductQuantity)]>> | 4th Quarter<<y [Where(c => c.Date.Month >= 10 && c.Date.Month <= 12).Sum(c => c.ProductQuantity)]>> |
| ---------- | ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ |
| Category 1 | 4.3                                                          | 2.4                                                          | 2                                                            | 3                                                            |
| Category 2 | 2.5                                                          | 4.4                                                          | 2                                                            | 2                                                            |
| Category 3 | 3.5                                                          | 1.8                                                          | 3                                                            | 5                                                            |
| Category 4 | 4.5                                                          | 2.8                                                          | 5                                                            | 2                                                            |

### Download Template

Please download the sample Chart Template we created in this article:

*   [Chart Template.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Email%20Templates/Chart%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_Dynamic_Title.msg)

### Generating The Report

{{< gist GroupDocsGists 0740b6ac3fefe47e4508be49b8939a4a >}}


