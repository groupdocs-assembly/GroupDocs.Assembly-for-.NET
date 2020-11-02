---
id: in-table-list-with-filtering-grouping-and-ordering-in-word-processing-document
url: assembly/net/in-table-list-with-filtering-grouping-and-ordering-in-word-processing-document
title: In-Table List with Filtering Grouping and Ordering in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-Table List with Filtering, Grouping, and Ordering report in Word Processing Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## In-Table List with Filtering, Grouping, and Ordering in Microsoft Word Document

### Creating a In-Table List with Filtering, Grouping, and Ordering

Practising the following steps you can create In-Table List with Filtering, Grouping, and Ordering Template in MS Word 2013.

1.  Click the document where you want to add the table.
2.  Press the "Insert" tab to insert the table.
3.  Insert a 2x2 table.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent customers' orders information with the following key requirements:

*   Report must show each customer along with sum of prices of his orders.
*   It must represent all the information in tabular form.
*   Report must be generated in the Word Processing Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="border-bottom-color: rgb(191, 191, 191); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(191, 191, 191); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(191, 191, 191); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(191, 191, 191); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 456.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> Price</span></p></td></tr><tr><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 266.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> [in </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">orders</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&nbsp;&nbsp;&nbsp; .</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Where</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">c.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">OrderDate</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Year</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> == </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">2015</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">)</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&nbsp;&nbsp;&nbsp; .</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">GroupBy</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">c.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">)</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&nbsp;&nbsp;&nbsp; .</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">OrderBy</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">(g =&gt; </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">g.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Key</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Name</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">)]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Key</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">.</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Name</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 167.55pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt;">Sum</span><span style="font-family: Calibri; font-size: 11pt;">(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.</span><span style="font-family: Calibri; font-size: 11pt;">Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;</span></p></td></tr></tbody></table><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&nbsp;</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download In-Table List with Filtering, Grouping, and Ordering Template

Please download the sample In-Table List with Filtering, Grouping, and Ordering document we created in this article:

*   [In-Table List with Filtering, Grouping, and Ordering.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering.docx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_DB.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_DB.docx?raw=true) (Template for DataSet, DataBase examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_XML.docx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_XML.docx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 1c97ce506e4b165e9bf131d26d943b8a >}}



#### Database Entities

{{< gist GroupDocsGists 1b4113a3fdc60a346e9fab3ece19ea3a >}}



#### Using DataSet

{{< gist GroupDocsGists 26fe1eeb21bb94d062f3fae206068cdd >}}



#### Using XML DataSource

{{< gist GroupDocsGists 95591cb01f8feeb8318415ba7bf18d4c >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 0ed1c013e8309329e1ee71569a179960 >}}



## In-Table List with Filtering, Grouping, and Ordering in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table List with Filtering, Grouping, and Ordering' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-table-reports-filtered-ordered-grouped/in-table-list-with-filtering-grouping-and-ordering-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table List with Filtering, Grouping, and Ordering\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_OpenDocument.odt?raw=true) (use same template for JSON example)
*   [In-Table List with Filtering, Grouping, and Ordering\_DB\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_DB_OpenDocument.odt?raw=true) (use same template for dataset examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_XML\_OpenDocument.odt](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_XML_OpenDocument.odt?raw=true)

### Generating the report

#### Custom Objects

{{< gist GroupDocsGists b70e818dbbcf0fa9577e8dd54734fdbe >}}



#### Database Entities

{{< gist GroupDocsGists 24899d1a3827a0d2b9fb0852d768983b >}}



#### Using DataSet

{{< gist GroupDocsGists 4e9929b8e6f24ecf73efb8f430f6e264 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 549f4b28bf37d8538d0d3d62c85edac8 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 6b3952f609dab8283974f13f8a931612 >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
