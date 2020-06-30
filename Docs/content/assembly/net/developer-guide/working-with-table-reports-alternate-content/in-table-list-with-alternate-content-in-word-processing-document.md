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
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Table List With Alternate Content report in Word Processing Document format based on the use case: Working with a Business Case.{{< /alert >}}

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

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="border-bottom-color: rgb(191, 191, 191); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(191, 191, 191); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(191, 191, 191); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(191, 191, 191); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 456.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Products</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> Price</span></p></td></tr><tr><td colspan="2" style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 445.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: center;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;if [!</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Any</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">()]&gt;&gt;No data</span></p></td></tr><tr><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;foreach [in </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">orders</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Product.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">ProductName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">P</span><span style="font-family: Calibri; font-size: 11pt;">rice</span><span style="font-family: Calibri; font-size: 11pt;">]&gt;&gt;&lt;&lt;/foreach&gt;&gt;</span></p></td></tr><tr><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Total:</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Sum</span><span style="font-family: Calibri; font-size: 11pt;">(c =&gt; c.</span><span style="font-family: Calibri; font-size: 11pt;">P</span><span style="font-family: Calibri; font-size: 11pt;">rice</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;&lt;&lt;/if&gt;&gt;</span></p></td></tr><tr style="height: 0pt;"><td style="width: 277.6pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 178.35pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td></tr></tbody></table>

{{< alert style="success" >}}For detailed technical information about syntax,expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download In-Table List With Alternate Content Template

Please download the sample In-Table List With Alternate Content document we created in this article:

*   [In-Table List With Alternate Content.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content.docx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List With Alternate Content\_DB.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_DB.docx?raw=true) (Template for DataBase examples)
*   [In-Table List With Alternate Content\_DS.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_DT.docx?raw=true) (Template for DataSet examples)
*   [In-Table List With Alternate Content\_XML.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Alternate%20Content_XML.docx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs f626dc3ae330f654a481 >}}



#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 5ff853b9ae1f50adeef5 >}}



#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 60dc72a869561606b285 >}}



#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 583e9e7adbdf2985ec71 >}}



#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs db90df496aee126c1db9 >}}



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

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs d4af481bdc81750096c1 >}}



#### Database Entities

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 0a05f8fbca8e3e226d49 >}}



#### Using DataSet

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 5e609e88739302d10f7e >}}



#### Using XML DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 59333326143ba5180bee >}}



#### Using JSON DataSource

{{< alert style="info" >}}The code uses some of the objects defined in: The Business Layer{{< /alert >}}{{< gist atirtahirgroupdocs 4cc6b9bdd03ef2daf75d >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
