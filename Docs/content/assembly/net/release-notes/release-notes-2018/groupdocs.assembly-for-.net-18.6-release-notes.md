---
id: groupdocs-assembly-for-net-18-6-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-6-release-notes
title: GroupDocs.Assembly for .NET 18.6 Release Notes
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.6.{{< /alert >}}

## Major Features

This release comes up with the support of null-conditional operators in template expressions and a number of improvements in working with charts.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-70  | Support dynamic coloring of chart series and individual series points for email messages with HTML and RTF bodies   | Feature |
| ASSEMBLYNET-76  | Support C# 6.0 null-conditional ?. and ?\[\] operators in template expressions  | Feature |
| ASSEMBLYNET-77  | Preserve the color of an of-pie chart split slice when using dynamic chart series coloring for Word Processing documents | Enhancement  |
| ASSEMBLYNET-78  | Preserve the font of a chart title while changing its text dynamically in emails  | Enhancement  |
| ASSEMBLYNET-79  | Preserve the static color of a chart series when dynamically filling it with data in emails  | Enhancement  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.6. It includes not only new and obsoleted public methods, but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly which may affect existing code. Any behaviour introduced that could be seen as a regression and modifies existing behaviour is especially important and is documented here.{{< /alert >}}

### Support dynamic coloring of chart series

For a chart with dynamic data, you can set colors of chart series dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series to be colored dynamically, define corresponding color expressions in names of these series using **seriesColor** tags having the following syntax:
    
    ```csharp
    <<seriesColor [color_expression]>>
    ```
    

### Support dynamic coloring of chart series points

For a chart with dynamic data, you can set colors of individual chart series points dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series with points to be colored dynamically, define corresponding color expressions in names of these series using **pointColor** tags having the following syntax:
    
    ```csharp
    <<pointColor [color_expression]>>
    ```
    

### Support C# 6.0 null-conditional ?. and ?\[\] operators in template expressions 

The following are the new operators that LINQ Reporting Engine enables you to use in template expressions:

*   x?.y  
*   a?\[x\]  
    

The engine follows operator precedence, associativity, and overload resolution rules declared at [C# Language Specification](http://www.microsoft.com/en-us/download/details.aspx?id=7029) while evaluating template expressions.
