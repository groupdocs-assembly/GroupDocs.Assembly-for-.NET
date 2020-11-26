---
id: groupdocs-assembly-for-net-20-1-release-notes
url: assembly/net/groupdocs-assembly-for-net-20-1-release-notes
title: GroupDocs.Assembly for .NET 20.1 Release Notes
weight: 100
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

Supported dynamic insertion of bookmarks for Word Processing documents and emails with HTML and RTF bodies and dynamic naming of cell ranges for Spreadsheet documents.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-139  | Support dynamic insertion of bookmarks for Word Processing documents  | Feature  |
| ASSEMBLYNET-140  | Support dynamic insertion of bookmarks for emails with HTML and RTF bodies  | Feature  |
| ASSEMBLYNET-141  | Support dynamic naming of cell ranges for Spreadsheet documents  | Feature  |
| ASSEMBLYNET-142  | Simple-type array with one element does not work for JsonDataSource  | Bug  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 20.1. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported dynamic insertion of bookmarks for Word Processing documents and emails with HTML and RTF bodies

You can insert bookmarks to your reports dynamically using bookmark tags. Syntax of a bookmark tag is defined as follows.

```csharp
<<bookmark [bookmark_expression]>>
bookmarked_content
<</bookmark>>
```

Here, bookmark\_expression defines the name of a bookmark to be inserted during runtime. This expression is mandatory and must return a non-empty value. While building a report, bookmark\_expression is evaluated and its result is used to construct a bookmark start and end that replace corresponding opening and closing bookmark tags respectively.

{{< alert style="warning" >}}A bookmark tag cannot be used within a chart.{{< /alert >}}

### Supported dynamic naming of cell ranges for Spreadsheet documents

Template syntax for dynamic naming of cell ranges for Spreadsheet documents is the same as above point. The difference is that it results to insertion of named cell ranges instead of bookmarks.
