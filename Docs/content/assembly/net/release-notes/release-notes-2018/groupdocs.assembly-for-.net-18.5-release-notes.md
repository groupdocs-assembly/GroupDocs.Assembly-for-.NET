---
id: groupdocs-assembly-for-net-18-5-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-5-release-notes
title: GroupDocs.Assembly for .NET 18.5 Release Notes
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.5.{{< /alert >}}

## Major Features

Support for dynamic coloring of chart series and individual series points for the most of supported file formats.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-67  | Support dynamic coloring of chart series and individual series points for Word Processing documents   | Feature |
| ASSEMBLYNET-68 | Support dynamic coloring of chart series and individual series points for Spreadsheet documents  | Feature |
| ASSEMBLYNET-69 | Support dynamic coloring of chart series and individual series points for Presentation documents  | Feature |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.5. It includes not only new and obsoleted public methods, but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly which may affect existing code. Any behaviour introduced that could be seen as a regression and modifies existing behaviour is especially important and is documented here.{{< /alert >}}

### Support dynamic coloring of chart series

For a chart with dynamic data, you can set colors of chart series dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series to be colored dynamically, define corresponding color expressions in names of these series using **seriesColor** tags having the following syntax:
    
    ```csharp
    <<seriesColor [color_expression]>>
    ```
    

### Support dynamic coloring of chart series points

For a chart with dynamic data, you can set colors of individual chart series points dynamically based upon expressions. To use the feature, do the following steps:

*   Declare a chart with dynamic data in the usual way
*   For chart series with points to be colored dynamically, define corresponding color expressions in names of these series using **pointColor** tags having the following syntax:
    
    ```csharp
    <<pointColor [color_expression]>>
    ```
