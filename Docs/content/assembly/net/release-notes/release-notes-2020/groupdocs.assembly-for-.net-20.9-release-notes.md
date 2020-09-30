---
id: groupdocs-assembly-for-net-20-9-release-notes
url: assembly/net/groupdocs-assembly-for-net-20-9-release-notes
title: GroupDocs.Assembly for .NET 20.9 Release Notes
weight: 50
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

GroupDocs.Assembly for .NET 20.9 extends working with Markdown by supporting tables and autolinks.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-160 | Support autolinks for Markdown | Feature |
| ASSEMBLYNET-161 | Support tables for Markdown | Feature |
| ASSEMBLYNET-164 | Adjust credit-based metered licensing to the latest policy | Feature |
| ASSEMBLYNET-162 | A number in a string is not correctly parsed for JsonDataSource | Bug |

## Public API and Backward Incompatible Changes 

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 20.9. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported tables and autolinks for Markdown

From now on, tables (see [Tables (extension)](https://github.github.com/gfm/#tables-extension-)) and autolinks (see chapter 6.7 of [Markdown specification](https://spec.commonmark.org/0.28/)) are supported when saving assembled Markdown documents to Word Processing formats and saving assembled Word Processing documents and emails to Markdown.
