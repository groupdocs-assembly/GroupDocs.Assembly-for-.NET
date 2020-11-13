---
id: multicolored-numbered-list-in-spreadsheet-document
url: assembly/net/multicolored-numbered-list-in-spreadsheet-document
title: Multicolored Numbered List in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Multicolored Numbered List report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Multicolored Numbered List in Microsoft Excel Document

### Creating a Multicolored Numbered List

Practising the following steps you can create Multicolored Numbered List Template in MS Excel 2013.

1.  Add a new Workbook.
2.  In one cell, write a sentence like "We provide support for the following products:".
3.  Select the cell, and assign a color by pressing "Cell Styles" in Styles group.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to describe the services you are providing with the following key requirements:

*   The report must show the products in the numbered list.
*   It must highlight the products.
*   The report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="border-bottom-color: rgb(191, 191, 191); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(191, 191, 191); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(191, 191, 191); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(191, 191, 191); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 456.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> Price</span></p></td></tr><tr><td style="background-color: rgb(255, 242, 204); border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;foreach [in </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">orders</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;&lt;&lt;if [</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Price</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> &gt;= </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">400</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Name</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(255, 242, 204); border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Price</span><span style="font-family: Calibri; font-size: 11pt;">]&gt;&gt;</span></p></td></tr><tr><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Name</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Price</span><span style="font-family: Calibri; font-size: 11pt;">]&gt;&gt;&lt;&lt;/if&gt;&gt;&lt;&lt;/foreach&gt;&gt;</span></p></td></tr><tr><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Sum</span><span style="font-family: Calibri; font-size: 11pt;">(c =&gt; c.</span><span style="font-family: Calibri; font-size: 11pt;">Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;</span></p></td></tr></tbody></table><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

Download Multicolored Numbered List Template

Please download the sample Multicolored Numbered List document we created in this article:

*   [Multicolored Numbered List.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Multicolored%20Numbered%20List.xlsx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Multicolored Numbered List\_XML.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Multicolored%20Numbered%20List_XML.xlsx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 85293a1d4e9e6576ed7a04b466da9bab >}}

#### Database Entities

{{< gist GroupDocsGists e6832a1c85f55ff178dc1501f94fbf17 >}}

#### Using DataSet

{{< gist GroupDocsGists 5ec824b55df1fca6151df7c619d77e30 >}}

#### Using XML DataSource

{{< gist GroupDocsGists 7d840e6770fc7c0993d154056ed8e4f4 >}}

#### Using JSON DataSource

{{< gist GroupDocsGists 7b8a5709e5ce0047bcfeb25027c2895b >}}

## Multicolored Numbered List in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Multicolored Numbered List' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Multicolored Numbered List\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Multicolored%20Numbered%20List_OpenDocument.ods?raw=true) (use same template for database, dataset and JSON examples)
*   [Multicolored Numbered List\_XML\_OpenDocument.ods](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Multicolored%20Numbered%20List_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists bd8efa56f534f00b2a58221e4ee14805 >}}

#### Database Entities

{{< gist GroupDocsGists afd28f07e88ce281f436be92f6ac9cef >}}

#### Using DataSet

{{< gist GroupDocsGists 04b66bd173782ebec7bfdb2f0cd72e2c >}}

#### Using XML DataSource

{{< gist GroupDocsGists 2f0f994bb60e119ea28665020d55e859 >}}

#### Using JSON DataSource

{{< gist GroupDocsGists 958e0d2af3d8ff265813aca35b37796d >}}

### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
