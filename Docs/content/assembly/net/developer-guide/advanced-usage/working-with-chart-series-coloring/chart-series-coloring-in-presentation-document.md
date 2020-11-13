---
id: chart-series-coloring-in-presentation-document
url: assembly/net/chart-series-coloring-in-presentation-document
title: Chart Series Coloring in Presentation Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.5 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Chart Series Coloring in Presentation Document

### Creating a Column Chart

Practising the following steps, you can insert a Column Chart in  MS PowerPoint 2013:

1.  Create a new presentation slide
2.  Click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box
3.  Select "Column" in the sidebar, you will see a gallery of charts
4.  Select the "100% Stacked Column" and press "OK" to insert the chart and Worksheet template to your document
5.  Edit the Worksheet with your data to update the chart. See [Chart Data (Excel)](https://docs.groupdocs.com/assembly/net/chart-series-coloring-in-presentation-document/#adding-syntax-to-be-evaluated-by-groupdocsassembly-engine).
6.  Save your Document

### Reporting Requirement

As a report developer, you are required to share contract price by manager dynamically with the following key requirements:

*   The report must show the  name of the manager
*   The report must show the total contract price for each manager 
*   Series color to be used in chart series 
*   The report must be generated in the Presentation Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
<Total Contract Prices by Managers<<foreach [m in
managers]>><<x [m.Manager]>>
```

#### Chart Data (Excel)

|   | Total Contract Price<<y [m.Total_Contract_Price]>><<seriesColor [color]>> |
| --- | --- |
| Category 1 | 4.3 |
| Category 2 | 2.5 |
| Category 3 | 3.5 |
| Category 4 | 4.5 |

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download Template

Please download the sample Dynamic Chart Series Color document we created in this article:

*   [Chart Template.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Dynamic%20Chart%20Series%20Color.pptx) (Template for CustomObject and JSON examples) 

### Generating The Report

For a chart with dynamic data, you can set colors of chart series dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series to be colored dynamically, define corresponding color expressions in names of these series using **seriesColor** tags having the following syntax:
  
    ```csharp
    <<seriesColor [color_expression]>>
    ```
    

A color expression must return a value of one of the following types:

*   A string containing the name of a known color, that is, the case-insensitive name of a member of the [KnownColor](https://msdn.microsoft.com/en-us/library/system.drawing.knowncolor(v=vs.110).aspx) enumeration such as "red"
*   An integer value defining RGB (red, green, blue) components of the color such as 0xFFFF00 (yellow)
*   A value of the [Color](http://msdn.microsoft.com/en-us/library/system.drawing.color(v=vs.110).aspx) type

Following code snippet generates the report: 

{{< gist GroupDocsGists de43d6f3fe81df28b528718fbceaff7d >}}


