---
id: groupdocs-assembly-for-net-19-1-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-1-release-notes
title: GroupDocs.Assembly for .NET 19.1 Release Notes
weight: 8
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.1.{{< /alert >}}

## Major Features

Supported dynamic merging of table cells with equal text contents.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-97  | Support dynamic merging of table cells containing the same text after template expressions are evaluated for Spreadsheet documents  | Feature  |
| ASSEMBLYNET-98  | Support dynamic merging of table cells containing the same text after template expressions are evaluated for Presentation documents  | Feature  |
| ASSEMBLYNET-104  | Support dynamic merging of table cells containing the same text after template expressions are evaluated for Word Processing documents  | Feature  |
| ASSEMBLYNET-105  | Support dynamic merging of table cells containing the same text after template expressions are evaluated for emails with HTML and RTF bodies  | Feature  |
| ASSEMBLYNET-106  | Support textual comments within template syntax tags  | Feature  |

  

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.1. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported dynamic merging of table cells with equal text contents.

You can merge table cells with equal textual contents within your reports dynamically using cellMerge tags. Syntax of a cellMerge tag is defined as follows.

```csharp
<<cellMerge -horz>>

```

A horz switch is optional. If the switch is present, it denotes a cell merging operation in a horizontal direction. If the switch is missing, it means that a cell merging operation is to be performed in a vertical direction (the default).

For two or more successive table cells to be merged dynamically in either direction by the engine, the following requirements must be met:

*   Each of the cells must contain a cellMerge tag denoting a cell merging operation in the same direction.
*   Each of the cells must not be already merged in another direction.
*   The cells must have equal textual contents (ignoring leading and trailing whitespaces).

Consider the following template.

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>&lt;&lt;cellMerge&gt;&gt;&lt;&lt;[value1]&gt;&gt;</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>&lt;&lt;cellMerge&gt;&gt;&lt;&lt;[value2]&gt;&gt;</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr></tbody></table>

If value1 and value2 have the same value, say "Hello", table cells containing cellMerge tags are successfully merged during runtime and a result report looks as follows then.

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td rowspan="2" class="confluenceTd"><p><strong>Hello</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr></tbody></table>

If value1 and value2 have different values, say "Hello" and "World", table cells containing cellMerge tags are not merged during runtime and a result report looks as follows then.

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>Hello</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>World</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr></tbody></table>

{{< alert style="warning" >}}A cellMerge tag can be normally used within a table data band.{{< /alert >}}

### Supported textual comments within template syntax tags

An optional comment providing a human-readable explanation ignored by the engine

```csharp
<<tag_name [expression] –switch1 –switch2 ... // optional_comment >>
```
