---
id: in-table-list-with-alternate-content-in-spreadsheet-document
url: assembly/net/in-table-list-with-alternate-content-in-spreadsheet-document
title: In-Table List With Alternate Content in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Table List With Alternate Content report in Spreadsheet Document format based on the use case: Working with a Business Case.{{< /alert >}}

## In-Table List With Alternate Content in Microsoft Excel Document

### Creating a In-Table List With Alternate Content

Practicing the following steps you can create In-Table List With Alternate Content Template in MS Excel 2013.

1.  Add a new Workbook.
2.  Select the range of cells that you want to include in the table.
3.  On the Insert tab, in the Tables group, click Table.
4.  Save your Document.

### Reporting Requirement

As a report developer, you are required to represent your products and their prices with the following key requirements:

*   Report must show each product along with its price.
*   It must show sum of all the prices.
*   It must represent all the information in tabular form.
*   Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 429pt;"><tbody><tr style="height: 15.75pt;"><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 274.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Products</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 132.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order Price</span></p></td></tr><tr style="height: 16.5pt;"><td colspan="2" style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(102, 102, 102); border-top-style: solid; border-top-width: 1.5pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 417.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: center;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;if [!Any()]&gt;&gt;No data</span></p></td></tr><tr style="height: 30.75pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 274.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> [in orders]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Product.ProductName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 132.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Price]&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;</span></p></td></tr><tr style="height: 30.75pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 274.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 132.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">$&lt;&lt;[Sum(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;&lt;&lt;/if&gt;&gt;</span></p></td></tr><tr style="height: 0pt;"><td style="width: 285pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 143pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download In-Table List With Alternate Content Template

Please download the sample In-Table List With Alternate Content document we created in this article:

*   [In-Table List With Alternate Content.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content.xlsx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List With Alternate Content\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_DB.xlsx?raw=true) (Template for DataBase examples)
*   [In-Table List With Alternate Content\_DS.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_DT.xlsx?raw=true) (Template for DataSet examples)
*   [In-Table List With Alternate Content\_XML.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_XML.xlsx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs a98766d8aa4cc2526f13 >}}



#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs cbc279906eec88a77084 >}}



#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 417279cba260aac6871a >}}



#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 489a120e45308c877256 >}}



#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 48d9770e11dea9add904 >}}



## In-Table List With Alternate Content in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table List With Alternate Content' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-table-reports-alternate-content/in-table-list-with-alternate-content-in-spreadsheet-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table List With Alternate Content\_OenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_OpenDocument.ods?raw=true) (use same template for JSON examples)
*   [In-Table List With Alternate Content\_DB\_OenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_DB_OpenDocument.ods?raw=true)
*   [In-Table List With Alternate Content\_DS\_OenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_DT_OpenDocument.ods?raw=true)
*   [In-Table List With Alternate Content\_XML\_OenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Alternate%20Content_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs e93a961dcf5400d1ada8 >}}



####   
Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 043c275824447b0009f7 >}}



#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 5e16dea1a139b5f45449 >}}



#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 9b785f2fb183b821c55f >}}



#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs dd6129229c8660473916 >}}



### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
