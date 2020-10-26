---
id: in-table-master-detail-in-email-document
url: assembly/net/in-table-master-detail-in-email-document
title: In-Table Master-Detail in Email Document
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate an In-Table Master-Detail report in Email Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

## In-Table Master-Detail in Email Document

### Creating a In-Table Master-Detail

Practising the following steps you can create In-Table Master-Detail Template in MS Outlook 2013.

1.  Create a new Email.
2.  Click the email where you want to add the table.
3.  Press "Insert" tab to insert the table.
4.  Insert a 2x4 table.
5.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent the following key requirements:

*   The report must be in .eml or .msg format.
*   It must add email recipient, CSS and subject of the email.
*   The report must show each customer along with his total orders prices.
*   It must also show each individual product ordered by the customers.
*   It must show sum of the order prices.
*   It must represent all the information in tabular form.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table class="tbl1">
	<tbody>
		<tr>
			<td colspan="2"> &lt;&lt;foreach [in customers]>> &lt;&lt;foreach [in Order]>> &lt;&lt;/foreach>> &lt;&lt;/foreach>></td>
		</tr>
		<tr>
			<td>Customer</td>
			<td>Order Price</td>
		</tr>
		<tr>
			<td><<[CustomerName]>>
</td>
			<td><<[Order.Sum(c => c.Price)]>></td>
		</tr>
		<tr>
			<td colspan="2">  <<[Product.ProductName]>> <<[Price]>></td>
		</tr>
		<tr>
			<td>Total:</td>
			<td><<[Sum( m => m.Order.Sum( c => c.Price))]>></td>
		</tr>
	</tbody>
</table>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download In-Table Master-Detail Template

Please download the sample In-Table Master-Detail document we created in this article:

*   [In-Table Master-Detail.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/In-Table%20Master-Detail.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists 000ce4e55c5a3f95e7e0fb87392c9e32 >}}


