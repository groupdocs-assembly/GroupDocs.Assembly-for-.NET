---
id: in-table-list-with-highlighted-rows-in-spreadsheet-document
url: assembly/net/in-table-list-with-highlighted-rows-in-spreadsheet-document
title: In-Table List with Highlighted Rows in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate an In-Table List with Highlighted Rows report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## In-Table List with Highlighted Rows in Microsoft Excel Document

### Creating a In-Table List with Highlighted Rows

Practising the following steps you can create In-Table List with Highlighted Rows Template in MS Excel 2013.

1.  Add a new Workbook.
2.  Select the range of cells that you want to include in the table.
3.  On the Insert tab, in the Tables group, click Table.
4.  Insert a 2x4 table.
5.  Click the cell you want to highlight.
6.  Click "Cell Styles" in Styles group.
7.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent customers' orders information with a specific filtering condition with the following key requirements:

*   The report must show each customer along with his orders.
*   Show single Customer and his single order price in a single row.
*   It must highlight the record with order price more than or equal to 400.
*   It must show sum of the order prices.
*   It must represent all the information in tabular form.
*   The report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 397pt;"><tbody><tr style="height: 15.75pt;"><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 231.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 143.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order Price</span></p></td></tr><tr style="height: 31.5pt;"><td style="background-color: rgb(255, 242, 204); border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 231.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> [in orders]&gt;&gt;&lt;&lt;if [Price &gt;= 400]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer.CustomerName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(255, 242, 204); border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 143.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Price]&gt;&gt;</span></p></td></tr><tr style="height: 15.75pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 231.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer.CustomerName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 143.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Price]&gt;&gt;&lt;&lt;/if&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;</span></p></td></tr><tr style="height: 15.75pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 231.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 143.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Sum(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download In-Table List with Highlighted Rows Template

Please download the sample In-Table List with Highlighted Rows document we created in this article:

*   [In-Table List with Highlighted Rows.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows.xlsx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List with Highlighted Rows\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_DB.xlsx?raw=true) (Template for DataBase examples)
*   [In-Table List with Highlighted Rows\_DT.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_DT.xlsx?raw=true) (Template for DataSet examples)
*   [In-Table List with Highlighted Rows\_XML.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_XML.xlsx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 9464bf4e1e33a4e3dcbfcab8cbe46787 >}}

#### Database Entities

{{< gist GroupDocsGists c372ab85a6ab05da66b315c61feef683 >}}

#### Using DataSet

{{< gist GroupDocsGists 3779bc487d0821d33aad53225ad0a095 >}}

#### Using XML DataSource

{{< gist GroupDocsGists dfa851238a40b6223d9394c5060e0f67 >}}

#### Using JSON DataSource

{{< gist GroupDocsGists 87fca83e4f81cb7720baf642bb2dadcf >}}

## In-Table List with Highlighted Rows in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table List with Highlighted Rows' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table List with Highlighted Rows\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_OpenDocument.ods?raw=true) (use same template for JSON examples)
*   [In-Table List with Highlighted Rows\_DB\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_DB_OpenDocument.ods?raw=true)
*   [In-Table List with Highlighted Rows\_DT\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_DT_OpenDocument.ods?raw=true)
*   [In-Table List with Highlighted Rows\_XML\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Highlighted%20Rows_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 0a078b84ed9ce47b4feb1a966b3518a4 >}}

#### Database Entities

{{< gist GroupDocsGists fed66320777b4940395fff94f4a3008e >}}

#### Using DataSet

{{< gist GroupDocsGists 246f0bc024abdc5410f41f229dacec12 >}}

#### Using XML DataSource

{{< gist GroupDocsGists b5d0c0e51ff601eac60fac7e0f1ed017 >}}

#### Using JSON DataSource

{{< gist GroupDocsGists 12df2bcb8cf6a3ba8da6983e8b10007b >}}

### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
