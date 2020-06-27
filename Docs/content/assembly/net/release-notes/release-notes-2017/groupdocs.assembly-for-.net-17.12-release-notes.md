---
id: groupdocs-assembly-for-net-17-12-release-notes
url: assembly/net/groupdocs-assembly-for-net-17-12-release-notes
title: GroupDocs.Assembly for .NET 17.12 Release Notes
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 17.12.{{< /alert >}}

## Major Features

This release of GroupDocs.Assembly extends abilities of dynamic manipulation over shapes, images, and charts for Word Processing and Email file formats.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-42  | Support dynamic insertion of images for MSG messages with RTF body   | Feature |
| ASSEMBLYNET-54 | Support dynamic shape fill color setting for Word Processing documents  | Feature |
| ASSEMBLYNET-55 | Support dynamic shape fill color setting for email messages with HTML body | Feature |
| ASSEMBLYNET-56 | Support dynamic shape fill color setting for email messages with RTF body | Feature |
| ASSEMBLYNET-57 | Support dynamic chart axis title setting for Word Processing documents  | Feature |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 17.12. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported dynamic chart axis title setting for Word Processing documents

From now, template syntax expressions are supported in chart axis titles. Thus, to set a chart axis title dynamically, you can put an expression tag into it like in the following example: 

```csharp
<<[axis_title_expression]>>
```

All common template syntax expression features are supported for chart axis titles with no restrictions.

### Support dynamic shape fill color setting for Word Processing and Email file formats

You can set text background color for document contents dynamically using backColor tags. The syntax of a backColor tag is defined as follows.

```csharp
<<backColor[color_expression]>>
content_to_be_colored
<</backColor>>

```

### Support dynamic insertion of images for MSG messages with RTF body 

Dynamic insertion of images and barcodes is now supported for MSG email messages with RTF body:

| Feature | MSG with RTF Body |
| --- | --- |
| Inserting Images Dynamically | Supported |
| Inserting Barcodes Dynamically | Supported |
