---
id: in-table-list-in-spreadsheet-document
url: assembly/net/in-table-list-in-spreadsheet-document
title: In-Table List in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-TableList report in Spreadsheet format based on the use case: Working with a Business Case.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer{{< /alert >}}

## In-Table List in Microsoft Excel Document

### Creating a In-Table List

Practicing the following steps you can create In-Table List Template in MS Excel 2013.

1.  Add a new Workbook.
2.  Select the range of cells that you want to include in the table.
3.  On the Insert tab, in the Tables group, click Table.
4.  Save your Document.

### Reporting Requirement

As a report developer, you are required to represent the information of the customers orders quantity with the following key requirements:

*   Report must show customers' name.
*   It must show the sum of orders prices against each customer.
*   It must sum up all the order prices for all the customers.
*   All the representation must be in tabular form.
*   Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 433pt;"><tbody><tr style="height: 15.75pt;"><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 259.95pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customers</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 150.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order Quantity</span></p></td></tr><tr style="height: 61.5pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 259.95pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> [in customers]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">CustomerName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 150.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Order.Sum</span><span style="font-family: Calibri; font-size: 11pt;">(</span><br><span style="font-family: Calibri; font-size: 11pt;">c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.ProductQuantity</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;</span></p></td></tr><tr style="height: 45.75pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 259.95pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 150.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Sum(</span><br><span style="font-family: Calibri; font-size: 11pt;">m =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">m.Order.Sum</span><span style="font-family: Calibri; font-size: 11pt;">(</span><br><span style="font-family: Calibri; font-size: 11pt;">c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.ProductQuantity</span><span style="font-family: Calibri; font-size: 11pt;">))]&gt;&gt;</span></p></td></tr></tbody></table>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download In-Table List Template

Please download the sample In-Table List document we created in this article:

*   [In-Table List.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List.xlsx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_DB.xlsx?raw=true) (Template for DataBase examples)
*   [In-Table List\_DS.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_DT.xlsx?raw=true) (Template for DataSet examples)
*   [In-Table List\_XML.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_XML.xlsx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 6813bfc47aea3f0c0963b05199b8b88d >}}



#### Database Entities

{{< gist GroupDocsGists b44f050de3d023c251f9b572c7589b81 >}}



#### Using DataSet

{{< gist GroupDocsGists b1cc321f36db86f679988b174bc735fb >}}



#### Using XML DataSource

{{< gist GroupDocsGists 451678be3fb64fb736646764c785a80e >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 6a5f78bf90c76876c2225bbd875348c5 >}}



## In-Table List in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table List' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-table-reports/in-table-list-in-spreadsheet-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table List\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_OpenDocument.ods?raw=true) (use same template for JSON examples)
*   [In-Table List\_DB\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_DB_OpenDocument.ods?raw=true)
*   [In-Table List\_DS\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_DT_OpenDocument.ods?raw=true)
*   [In-Table List\_XML\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists c2343859b5d4894fff20b160ee0f16d1 >}}



#### Database Entities

{{< gist GroupDocsGists fecaeddcf4e9a99820e0cdbf782cc852 >}}



#### Using DataSet

{{< gist GroupDocsGists 508487e8a09c41d1b318fb8757ec3301 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 57b27442f9843eb170a024da982d9213 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists f7056c8c0d2a25ced13b40f80201c2fd >}}



### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
