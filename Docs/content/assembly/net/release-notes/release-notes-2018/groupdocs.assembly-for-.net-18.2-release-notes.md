---
id: groupdocs-assembly-for-net-18-2-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-2-release-notes
title: GroupDocs.Assembly for .NET 18.2 Release Notes
weight: 7
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.2.{{< /alert >}}

## Major Features

This release extends abilities to work with single-column table data bands and conditional blocks.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-61  | Support single-column table-row data bands for Word Processing documents     | Feature |
| ASSEMBLYNET-62 | Support single-column table-row data bands for Spreadsheet documents | Feature |
| ASSEMBLYNET-63 | Support single-column table-row data bands for Presentation documents  | Feature |
| ASSEMBLYNE4\-63 | Support single-column table-row data bands for email messages with HTML and RTF bodies | Feature  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.2. It includes not only new and obsoleted public methods, but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly which may affect existing code. Any behaviour introduced that could be seen as a regression and modifies existing behaviour is especially important and is documented here.{{< /alert >}}

### Working with Table-Row Data Bands

A table-row data band is a data band which body occupies single or multiple rows of a single document table. The body of such a band starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows.

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;foreach ...&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>... &lt;&lt;/foreach&gt;&gt;</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td></tr></tbody></table>

### Working with Table-Row Conditional Blocks

A table-row conditional block is a conditional block which body occupies single or multiple rows of a single document table. The body of such a block (as well as the body of its every template option) starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows.

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;if ...&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;elseif ...&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;else&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>... &lt;&lt;/if&gt;&gt;</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td></tr></tbody></table>
