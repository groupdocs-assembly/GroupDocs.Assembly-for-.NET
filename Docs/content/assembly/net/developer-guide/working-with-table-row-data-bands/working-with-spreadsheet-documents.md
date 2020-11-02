---
id: working-with-spreadsheet-documents
url: assembly/net/working-with-spreadsheet-documents
title: Working with Spreadsheet Documents
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.2 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Working with Spreadsheet Documents

GroupDocs.Assembly allows you to use data bands in table rows in **Spreadsheet Documents**. A table-row data band is a data band which body occupies single or multiple rows of a single document table. The body of such a band starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows. The syntax for using data bands in spreadsheet document is as follows:

<table class="gd-assembly">
	<tbody>
		<tr>
			<td>Manager</td>
			<td>Manager</td>
			<td>Contract Price</td>
		</tr>
		<tr>
			<td>&lt;&lt;foreach [c in ds]>>&lt;&lt;[c.Client]>></td>
			<td>&lt;&lt;[c.Manager]>></td>
			<td>&lt;&lt;[c.Price]>>&lt;&lt;/foreach>></td>
		</tr>
	</tbody>
</table>

### Working with Table-Row Conditional Blocks

You can also use a table-row conditional block which body occupies single or multiple rows of a single document table. The body of such a block (as well as the body of its every template option) starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows:

|                        |         |                 |
| ---------------------- | ------- | --------------- |
| **&lt;&lt;if ...>> ...**     | **...** | **...**         |
| **...**                | **...** | **...**         |
| **&lt;&lt;elseif ...>> ...** | **...** | **...**         |
| **...**                | **...** | **...**         |
| **&lt;&lt;else>> ...**       | **...** | **...**         |
| **...**                | **...** | **...**         |
| **...**                | **...** | **... &lt;&lt;/if>>** |
|                        |         |                 |

### Download

#### Data Source Document

*   [Managers Data.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/Excel%20DataSource/Contracts%20Data.xlsx)

#### Template

*   [Working with Table Row Data Bands.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Working%20With%20Table%20Row%20Data%20Bands.xlsx)

### The Code

{{< gist GroupDocsGists 4fbdcf4fde19831fbb9589f1d2615611 >}}


