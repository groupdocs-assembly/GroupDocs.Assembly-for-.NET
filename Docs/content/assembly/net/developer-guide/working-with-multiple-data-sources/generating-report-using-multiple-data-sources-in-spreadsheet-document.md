---
id: generating-report-using-multiple-data-sources-in-spreadsheet-document
url: assembly/net/generating-report-using-multiple-data-sources-in-spreadsheet-document
title: Generating Report using Multiple Data Sources in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a report. This report will fetch data from multiple data sources.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Generating Report in Microsoft Excel Document

### Creating a Template

1.  Insert a desired shape to display image in it.
2.  Go to Insert Tab and select shape by clicking on Shape Icon.
3.  Add a bullet list at the place where you want it
4.  Save your Document

### Reporting Requirement

As a report developer, you are required to generate a report that fetches data from two different data sources (e:g XML, JSON). Report must show following information:

*   Customer name and picture
*   List of products

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table class="gd-assembly">
	<tbody>
		<tr>
			<td colspan="3">&lt;&lt;foreach [in customers]>></td>
			<td></td>
			<td></td>
		</tr>
		<tr style="vertical-align: top">
			<td colspan="3" rowspan="8">&lt;&lt;image[customer.Photo]>></td>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td colspan="3">&lt;&lt;[CustomerName]>></td>
			<td colspan="2">&lt;&lt;/foreach>></td>
		</tr>
	</tbody>
</table>

**In-Cell List**

```csharp
We provide support for the following products:
<<foreach [in products]>>
-          <<[ProductName]>><</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download Template

Get template from here.

*   [Multiple DS.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Multiple%20DS.xlsx?raw=true)

### Generating the Report

{{< gist atirtahirgroupdocs 912c6398a7f6639c57447241b674a5ff >}}



## Generating Report in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating a report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-multiple-data-sources/generating-report-using-multiple-data-sources-in-spreadsheet-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

Get template from here.

*   [Multiple DS.ods](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Multiple%20DS.ods?raw=true)

### Generating the Report

{{< gist GroupDocsGists 9a11804cc2ada59aa5e84776c5dbb2c5 >}}


