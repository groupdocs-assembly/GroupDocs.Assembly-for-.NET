---
id: common-master-detail-image-in-spreadsheet-document
url: assembly/net/common-master-detail-image-in-spreadsheet-document
title: Common Master-Detail Image in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail Image report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Common Master-Detail Image in Microsoft Excel Document

### Creating a Common Master-Detail Image

Please follow below steps to create Common Master-Detail Image Template in MS Excel 2013:

1.  Create a new Workbook.
2.  Insert a shape to display image in it.
3.  Go to Insert Tab and select shape by clicking on Shape Icon.
4.  Save the document.

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   Report must show customers' picture and name.
*   It must associate the customers with their products.
*   Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

&lt;&lt;foreach \[in customers\]>>

<table class="gd-assembly1"cellspacing="0" cellpadding="0" style="border-collapse: collapse; float: bottom; margin-top: 0pt; margin-right: 9pt; margin-bottom: 0pt; margin-left: 9pt; width: 132.15pt; height: 132.15pt ">
	<tbody>
		<tr>
			<td style="vertical-align: top;">&lt;&lt;image [customer.Photo]>></td>
		</tr>
	</tbody>
</table>

**&lt;&lt;[CustomerName]>>**

**Clients: &lt;&lt;foreach [in Order]>>&lt;&lt;[IndexOf() != 0 ? ", " : ""]>>&lt;&lt;[Product.ProductName]>>&lt;&lt;/foreach>>**

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Common Master-DetailTemplate

Please download the sample Common Master-Detail document we created in this article:

*   [Common Master-Detail.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Common%20Master-Detail.xlsx?raw=true) (Template for CustomObject and JSON examples)
*   [Common Master-Detail\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Common%20Master-Detail_DB.xlsx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists d975e0dedc5f473afb719f3775cc6c8e >}}



#### Database Entities

{{< gist GroupDocsGists fcc7a15dfd4db8f43f0e6110dd34c8c8 >}}



#### Using DataSet

{{< gist GroupDocsGists 4207c6e947d0ddd568e1a245f33a0990 >}}



#### Using XML DataSource

{{< gist GroupDocsGists fad61412f602a116a5b473d46e659c67 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 76715602bb202e81dcb316d1d58b9949 >}}



## Common Master-Detail in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit wikipedia article.

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Common Master-Detail' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-image-reports-common-master-detail/common-master-detail-image-in-spreadsheet-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Common Master-Detail\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Common%20Master-Detail_OpenDocument.ods?raw=true) (use same template for JSON examples)
*   [Common Master-Detail\_DB\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Common%20Master-Detail_DB_OpenDocument.ods?raw=true) (use same template for database and dataset and XML examples)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 10f4b7f883ebd432815e3069527736a6 >}}



#### Database Entities

{{< gist GroupDocsGists fb728a436de90b71847f7fc5b5e3d7fa >}}



#### Using DataSet

{{< gist GroupDocsGists 05ac56a212793baa1eb02726a416483e >}}



#### Using XML DataSource

{{< gist GroupDocsGists f9e1ee9f348b65659cb3d6b3882618a9 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 751ec891d4d35bd044f0d55c904ee60e >}}



### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
