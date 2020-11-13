---
id: groupdocs-assembly-for-net-19-12-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-12-release-notes
title: GroupDocs.Assembly for .NET 19.12 Release Notes
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.12{{< /alert >}}

### Major Features

Supported unordered lists for Markdown and access to related DataTable using relation name.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-135  | Support unordered lists for Markdown  | Feature  |
| ASSEMBLYNET-136  | Support access to related DataTable using relation name  | Feature  |
| ASSEMBLYNET-137  | NullReferenceException is thrown when evaluating a null-conditional operator on missing related DataRow  | Bug  |
| ASSEMBLYNET-138  | Page breaks are removed if DocumentAssemblyOptions.RemoveEmptyParagraphs is used  | Bug  |

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.12. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

## Public API and Backward Incompatible Changes 

### Supported access to related DataTable using relation name

The Document Assembly Engine enables you to access a data associated with a particular *DataRow* or *DataRowView* instance in template expressions using the "." operator. See "[Working with DataTable and DataView Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information.

### Supported unordered lists for Markdown

From now on, unordered lists (see chapter 5.3 of [Markdown specification](https://spec.commonmark.org/0.28/)) are supported when saving assembled Markdown documents to Word Processing formats and emails to Markdown.
