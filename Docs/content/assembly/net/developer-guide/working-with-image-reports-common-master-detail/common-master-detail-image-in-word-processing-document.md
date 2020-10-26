---
id: common-master-detail-image-in-word-processing-document
url: assembly/net/common-master-detail-image-in-word-processing-document
title: Common Master-Detail Image in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail Image report in Word Processing Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Common Master-Detail Image in Microsoft Word Document

### Creating a Common Master-Detail Image

Please follow below steps you can create Common Master-Detail Template in MS Word 2013.

1.  Insert the desired shape to display image in it.
2.  Go to Insert Tab and select shape by clicking on Shape Icon.
3.  Save your Document.

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   Report must show customers' picture and name.
*   It must associate the customers with their products.
*   Report must be generated in the Word Processing Document.  
      
  

{{< alert style="info" >}}See how to use images in MS Word here.{{< /alert >}}

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

**Products: &lt;&lt;foreach [in Order]>>&lt;&lt;[IndexOf() != 0 ? ", " : ""]>>&lt;&lt;[Product.ProductName]>>&lt;&lt;/foreach>>**

&lt;&lt;/foreach>>

{{< alert style="success" >}}

For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).

{{< /alert >}}

### Download Common Master-DetailTemplate

Please download the sample Common Master-Detail document we created in this article:

*   [Common Master-Detail.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail.docx?raw=true) (Template for CustomObject and JSON examples)
*   [Common Master-Detail\_DB.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail_DB.docx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists f659589cae0f24c5500f9e2fd39c8ff5 >}}



#### Database Entities

{{< gist GroupDocsGists 0f0827a650b0b0e4db55ddca1efa8033 >}}



#### Using DataSet

{{< gist GroupDocsGists 004ac14e5a96b877eefe0b42a7e6e5ef >}}



#### Using XML DataSource

{{< gist GroupDocsGists a8303e24fefbd71ceff390a32e13668b >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 721f9f83f25c880bc9f57383f17d0d6e >}}



## Common Master-Detail in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Common Master-Detail' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-image-reports-common-master-detail/common-master-detail-image-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Common Master-Detail\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail_OpenDocument.odt?raw=true) (use same template for JSON examples)
*   [Common Master-Detail\_DB\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail_DB_OpenDocument.odt?raw=true) (use same template for database and dataset and XML examples)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 9e3a02f4a1f5b7ef54ccb059dd96a039 >}}



#### Database Entities

{{< gist GroupDocsGists b271cec9679d51438f4990085f37ba02 >}}



#### Using DataSet

{{< gist GroupDocsGists 3d522b9b2cae2e11647d4089a4f2dfd7 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 4a676c7c9c231e5fc4f489d36aa8d370 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 352b28ee4af2e26ca54a188d5bdbfd16 >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
