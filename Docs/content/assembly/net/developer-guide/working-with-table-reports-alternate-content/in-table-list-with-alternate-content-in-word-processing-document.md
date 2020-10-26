---
id: in-table-list-with-alternate-content-in-word-processing-document
url: assembly/net/in-table-list-with-alternate-content-in-word-processing-document
title: In-Table List With Alternate Content in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Table List With Alternate Content report in Word Processing Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}

## In-Table List With Alternate Content in Microsoft Word Document

### Creating a In-Table List With Alternate Content

Practicing the following steps you can create In-Table List With Alternate Content Template in MS Word 2013.

1.  Click the document where you want to add the table.
2.  Press "Insert" tab to insert the table.
3.  Insert a 2X3 table.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent your products and their prices with the following key requirements:

*   Report must show each product along with its price.
*   It must show sum of all the prices.
*   It must represent all the information in tabular form.
*   Report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table class="gd-assembly">
	<tbody>
		<tr>
			<td>Products</td>
			<td>Order Price</td>
		</tr>
		<tr>
			<td colspan="2">&lt;&lt;if [!Any()]>>No data</td>
		</tr>
		<tr>
			<td>&lt;&lt;else>>&lt;&lt;foreach [in orders]>><br>&lt;&lt;[Product.ProductName]>></td>
			<td>&lt;&lt;[Price]>>&lt;&lt;/foreach>></td>
		</tr>
		<tr>
			<td>Total:</td>
			<td>&lt;&lt;[Sum(c => c.Price)]>>&lt;&lt;/if>></td>
		</tr>
	</tbody>
</table>

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download In-Table List With Alternate Content Template

Please download the sample In-Table List With Alternate Content document we created in this article:

*   [In-Table List With Alternate Content.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content.docx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List With Alternate Content\_DB.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_DB.docx?raw=true) (Template for DataBase examples)
*   [In-Table List With Alternate Content\_DS.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_DT.docx?raw=true) (Template for DataSet examples)
*   [In-Table List With Alternate Content\_XML.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_XML.docx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs f626dc3ae330f654a481 >}}



#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 5ff853b9ae1f50adeef5 >}}



#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 60dc72a869561606b285 >}}



#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 583e9e7adbdf2985ec71 >}}



#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs db90df496aee126c1db9 >}}



## In-Table List With Alternate Content in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table List with Alternate Content' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-table-reports-alternate-content/in-table-list-with-alternate-content-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table List With Alternate Content\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_OpenDocument.odt?raw=true) (use same template for JSON examples)
*   [In-Table List With Alternate Content\_DB\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_DB_OpenDocument.odt?raw=true)
*   [In-Table List With Alternate Content\_DS\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_DT_OpenDocument.odt?raw=true)
*   [In-Table List With Alternate Content\_XML\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_XML_OpenDocument.odt?raw=true)

### Generating the report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs d4af481bdc81750096c1 >}}



#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 0a05f8fbca8e3e226d49 >}}



#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 5e609e88739302d10f7e >}}



#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 59333326143ba5180bee >}}



#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}{{< gist atirtahirgroupdocs 4cc6b9bdd03ef2daf75d >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
