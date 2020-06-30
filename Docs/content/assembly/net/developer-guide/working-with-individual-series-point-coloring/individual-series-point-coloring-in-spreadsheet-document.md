---
id: individual-series-point-coloring-in-spreadsheet-document
url: assembly/net/individual-series-point-coloring-in-spreadsheet-document
title: Individual Series Point Coloring in Spreadsheet Document
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.5 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Individual Series Point Coloring in Spreadsheet Document

Please follow below steps to can create Pie Chart in MS Excel 2013.

1.  Add a new Workbook
2.  Click on the workbook where you want to insert the chart, click the "Insert" tab, and then click "Pie Chart Icon" in the charts group
3.  A drop-down with charts will appear, select the "Pie" and press "OK" to insert the chart
4.  Click on the chart you just inserted, then click the "Change Data" icon in Data group
5.  Now add legend entries. See [Chart Data](https://docs.dynabic.com/display/assemblynet/Pie+Chart+in+Spreadsheet+Document#PieChartinSpreadsheetDocument-ChartData)
6.  Save your Document

### Reporting Requirement

As a report developer, you are required to share your customers' orders details dynamically with the following key requirements:

*   The report must show information on a Pie Chart
*   It must indicate customer name with value (price of the orders purchased)
*   The report must be generated in the Spreadsheet Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Price<<foreach [in customers]>>
<<x [CustomerName]>>

```

####   
Chart Data

**Legend Entries**

```csharp
="Total Order Price<<y [Order.Sum(c => c.Price)]>><<pointColor[color]>>"
```

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Pie Chart Template

Please download the sample Pie Chart document we created in this article:

*   [Dynamic Chart Point Series Color.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Dynamic%20Chart%20Point%20Series%20Color.xlsx) (Template for CustomObject and JSON examples) 

### Generating The Report

For a chart with dynamic data, you can set colors of individual chart series points dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series with points to be colored dynamically, define corresponding color expressions in names of these series using **pointColor** tags having the following syntax:
    
    ```csharp
    <<pointColor [color_expression]>>
    ```
    

A color expression must return a value of one of the following types:

*   A string containing the name of a known color, that is, the case-insensitive name of a member of the [KnownColor](https://msdn.microsoft.com/en-us/library/system.drawing.knowncolor(v=vs.110).aspx) enumeration such as "red"
*   An integer value defining RGB (red, green, blue) components of the color such as 0xFFFF00 (yellow)
*   A value of the [Color](http://msdn.microsoft.com/en-us/library/system.drawing.color(v=vs.110).aspx) type

Following code snippet generates the report:

{{< gist GroupDocsGists 0d47085c93fbc137e03d6f586f3fecd4 >}}


