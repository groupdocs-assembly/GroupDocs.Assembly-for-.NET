---
id: generating-report-using-multiple-data-sources-in-word-processing-document
url: assembly/net/generating-report-using-multiple-data-sources-in-word-processing-document
title: Generating Report using Multiple Data Sources in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a report. This report will fetch data from multiple data sources.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Generating Report in Microsoft Word Document

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

&lt;&lt;foreach [in customers]>>

<table class="gd-assembly1"cellspacing="0" cellpadding="0" style="border-collapse: collapse; float: bottom; margin-top: 0pt; margin-right: 9pt; margin-bottom: 0pt; margin-left: 9pt; width: 132.15pt; height: 132.15pt ">
	<tbody>
		<tr>
			<td style="vertical-align: top;">&lt;&lt;image [customer.Photo]>></td>
		</tr>
	</tbody>
</table>

**&lt;&lt;[CustomerName]>>**

&lt;&lt;/foreach>>

**List of Products**

```csharp
We provide support for the following products:
. <<foreach [in products]>><<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download Template

Get template from here.

*   [Multiple DS.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Multiple%20DS.docx?raw=true)

### Generating the Report

{{< gist GroupDocsGists c4e760407e1a84ba8ca863cf96faede0 >}}



## Generating Report in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating a report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-multiple-data-sources/generating-report-using-multiple-data-sources-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

Get template from here.

*   [Multiple DS.odt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Multiple%20DS.odt?raw=true)

### Generating the Report

{{< gist GroupDocsGists b2875d58468c1f71d70f4f9bc1cec833 >}}


