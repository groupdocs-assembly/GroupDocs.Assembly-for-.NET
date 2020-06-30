---
id: working-with-presentation-documents
url: assembly/net/working-with-presentation-documents
title: Working with Presentation Documents
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.2 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Working with Presentation Documents

GroupDocs.Assembly allows you to use data bands in table rows in **Presentation**** Documents**. A table-row data band is a data band which body occupies single or multiple rows of a single document table. The body of such a band starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows. The syntax for using data bands in presentation document is as follows:

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>Client</strong></p></td><td class="confluenceTd"><p><strong>Manager</strong></p></td><td class="confluenceTd"><p><strong>Contract Price</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;foreach [</strong></p><p><strong>&nbsp;&nbsp;&nbsp; c in ds</strong></p><p><strong>]&gt;&gt;&lt;&lt;[c.Client]&gt;&gt;</strong></p></td><td class="confluenceTd"><p><strong>&lt;&lt;[c.Manager]&gt;&gt;</strong></p></td><td class="confluenceTd"><p><strong>&lt;&lt;[c.Price]&gt;&gt;&lt;&lt;/</strong></p><p><strong>foreach&gt;&gt;</strong></p></td></tr></tbody></table>

### Working with Table-Row Conditional Blocks

You can also use a table-row conditional block which body occupies single or multiple rows of a single document table. The body of such a block (as well as the body of its every template option) starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows:

<table class="confluenceTable"><tbody><tr><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;if ...&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;elseif ...&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&lt;&lt;else&gt;&gt; ...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>...</strong></p></td><td class="confluenceTd"><p><strong>... &lt;&lt;/if&gt;&gt;</strong></p></td></tr><tr><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p></td><td class="confluenceTd"><p><strong>&nbsp;</strong></p><div><strong><br></strong></div></td></tr></tbody></table>

### Download

#### Data Source Document

*   [Managers Data.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Excel%20DataSource/Contracts%20Data.xlsx)

#### Template

*   [Working with Table Row Data Bands.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Working%20With%20Table%20Row%20Data%20Bands.pptx)

### The Code

{{< gist GroupDocsGists 89c648653a7515892dc25e42c94b56e5 >}}


