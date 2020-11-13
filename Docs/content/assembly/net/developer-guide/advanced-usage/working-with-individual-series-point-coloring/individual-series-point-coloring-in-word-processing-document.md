---
id: individual-series-point-coloring-in-word-processing-document
url: assembly/net/individual-series-point-coloring-in-word-processing-document
title: Individual Series Point Coloring in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.5 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Individual Series Point Coloring in Word Processing Document

Please follow below steps to can create Pie Chart in MS Word 2013.

1.  Click in the document where you want to insert the chart, click the "Insert" tab, and then click "Chart" in the illustrations group to open the "Insert Chart" dialog box.
2.  Select "Pie".
3.  Preview "Pie" and press OK to insert the chart and Worksheet template to your document.
4.  Edit the Worksheet with your data to update the chart. See [Chart Data (Excel)](https://docs.groupdocs.com/assembly/net/individual-series-point-coloring-in-word-processing-document/#adding-syntax-to-be-evaluated-by-groupdocsassembly-engine)
5.  Save the template.

### Reporting Requirement

As a report developer, you are required to share your customers' orders details dynamically with the following key requirements:

*   The report must show information on a Pie Chart
*   It must indicate customer name with value (price of the orders purchased)
*   The report must be generated in the Word Processing Document

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

#### Chart Title

```csharp
Total Order Price<<foreach [in customers]>>
<<x [CustomerName]>>
```

#### Chart Data (Excel)

|   | Total Order Price<<y [Order.Sum(c => c.Price)]>><<pointColor [Color] >> |
| --- | --- |
| A | 8.2 |
| B | 3.2 |
| C | 1.5 |
| D | 1.2 |

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download Pie Chart Template

Please download the sample Pie Chart document we created in this article:

*   [Dynamic Chart Point Series Color.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Dynamic%20Chart%20Point%20Series%20Color.docx) (Template for CustomObject and JSON examples) 

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

{{< gist GroupDocsGists fb1400ea59b283001465944f5a2780d8 >}}


