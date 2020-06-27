---
id: groupdocs-assembly-for-net-18-1-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-1-release-notes
title: GroupDocs.Assembly for .NET 18.1 Release Notes
weight: 8
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.1.{{< /alert >}}

## Major Features

Extended support of dynamic chart axis title setting for more file formats.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-58  | Support dynamic chart axis title setting for Spreadsheet documents    | Feature |
| ASSEMBLYNET-59 | Support dynamic chart axis title setting for Presentation documents  | Feature |
| ASSEMBLYNET-60 | Support dynamic chart axis title setting for emails with HTML and RTF bodies | Feature |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.1. It includes not only new and obsoleted public methods, but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly which may affect existing code. Any behaviour introduced that could be seen as a regression and modifies existing behaviour is especially important and is documented here.{{< /alert >}}

### Supported dynamic chart axis title setting for Spreadsheet, Presentation, and Email file formats

From now, template syntax expressions are supported in chart axis titles. Thus, to set a chart axis title dynamically, you can put an expression tag into it like in the following example: 

```csharp
<<[axis_title_expression]>>
```

All common template syntax expression features are supported for chart axis titles with no restrictions.
