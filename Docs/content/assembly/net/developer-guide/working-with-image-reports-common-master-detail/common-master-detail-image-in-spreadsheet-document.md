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
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail Image report in Spreadsheet Document format based on the use case: Working with a Business Case.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

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

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 722.9pt;"><tbody><tr style="height: 15pt;"><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: middle; width: 142pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;"> [in customers]&gt;&gt;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td><td style="vertical-align: top;"></td><td style="vertical-align: top;"></td><td style="vertical-align: top;"></td></tr><tr style="height: 133.45pt;"><td colspan="12" style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 678.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr style="height: 119.25pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 131.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="background-color: rgb(255, 255, 255); color: rgb(51, 51, 51); font-family: Calibri; font-size: 11pt;">&lt;&lt;image [customer.Photo]&gt;&gt;</span></p></td></tr><tr style="height: 15pt;"><td style="vertical-align: bottom; width: 142pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td></tr></tbody></table><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"></p></td><td style="vertical-align: top;"></td><td style="vertical-align: top;"></td><td style="vertical-align: top;"></td></tr><tr style="height: 21pt;"><td colspan="5" style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: middle; width: 337.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">CustomerName</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 38pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 0.3pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 0.3pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td><td style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 0.3pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt;">&nbsp;</span></p></td></tr><tr style="height: 21pt;"><td colspan="15" style="padding-left: 5.4pt; padding-right: 5.4pt; vertical-align: bottom; width: 712.1pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">Clients: &lt;&lt;</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;"> [in Order]&gt;&gt;&lt;</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">&lt;[</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">IndexOf</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">() != 0 ? ", " : ""]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">Product.ProductName</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">]&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 12pt; font-weight: bold;">&gt;&gt;</span></p></td></tr><tr style="height: 0pt;"><td style="width: 152.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 48.8pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 11.1pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 11.1pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 11.1pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

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
