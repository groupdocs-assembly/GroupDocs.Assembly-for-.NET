---
id: groupdocs-assembly-for-net-21-4-release-notes
url: assembly/net/groupdocs-assembly-for-net-21-4-release-notes
title: GroupDocs.Assembly for .NET 21.4 Release Notes
weight: 80
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

Provided dynamic cell merging in both directions simultaneously and supported Select and SelectMany LINQ extension methods in template syntax.

## Full List of Features Covering all Changes in this Release

| Key             | Summary                                                      | Category |
| :-------------- | :----------------------------------------------------------- | :------- |
| ASSEMBLYNET-183 | Support dynamic cell merging in both directions simultaneously for Word Processing documents | Feature  |
| ASSEMBLYNET-184 | Support dynamic cell merging in both directions simultaneously for emails with HTML and RTF bodies | Feature  |
| ASSEMBLYNET-185 | Support dynamic cell merging in both directions simultaneously for Spreadsheet documents | Feature  |
| ASSEMBLYNET-186 | Support dynamic cell merging in both directions simultaneously for Presentation documents | Feature  |
| ASSEMBLYNET-193 | Support SelectMany LINQ extension method in template syntax  | Feature  |
| ASSEMBLYNET-194 | Support Select LINQ extension method in template syntax      | Feature  |
| ASSEMBLYNET-195 | Improve loading of integer values from JSON, XML, and CSV    | Bug      |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 21.4. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Improved Loading of Integer Values from JSON, XML, and CSV

{{< alert style="warning" >}}

The feature is available for all supported file formats.

{{< /alert >}}

A detailed description of this feature can be found in the ["Accessing XML Data"](https://docs.groupdocs.com/assembly/net/groupdocs-assembly-engine-apis/#accessing-xml-data), ["Accessing JSON Data"](https://docs.groupdocs.com/assembly/net/groupdocs-assembly-engine-apis/#accessing-json-data), and ["Accessing CSV Data"](https://docs.groupdocs.com/assembly/net/groupdocs-assembly-engine-apis/#accessing-csv-data) sections.

### Supported Dynamic Cell Merging in Both Directions Simultaneously

{{< alert style="warning" >}}

The feature is available for Word Processing, Spreadsheet, and Presentation documents, as well as for emails with HTML and RTF bodies.

{{< /alert >}}

A detailed description of this feature can be found in the ["Merging Table Cells Dynamically"](https://docs.groupdocs.com/assembly/net/template-syntax-part-2-of-2/#merging-table-cells-dynamically) section.

### Supported Select and SelectMany LINQ Extension Methods in Template Syntax

{{< alert style="warning" >}}

The feature is available for all supported file formats.

{{< /alert >}}

A detailed description of this feature can be found in the ["Enumeration Extension Methods"](https://docs.groupdocs.com/assembly/net/enumeration-extension-methods) section.