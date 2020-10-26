---
id: in-table-list-with-filtering-grouping-and-ordering-in-spreadsheet-document
url: assembly/net/in-table-list-with-filtering-grouping-and-ordering-in-spreadsheet-document
title: In-Table List with Filtering Grouping and Ordering in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
# In-Table List with Filtering, Grouping, and Ordering in Spreadsheet Document

{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a In-TableList with Filtering, Grouping, and Ordering report in Spreadsheet Document format based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## In-Table List with Filtering, Grouping, and Ordering in Microsoft Excel Document

### Creating a In-Table List with Filtering, Grouping, and Ordering

Practising the following steps you can create In-Table List with Filtering, Grouping, and Ordering Template in MS Excel 2013.

1.  Add a new Workbook.
2.  Select the range of cells that you want to include in the table.
3.  On the Insert tab, in the Tables group, click Table.
4.  Save your Document.

### Reporting Requirement

As a report developer, you are required to represent customers' orders information with the following key requirements:

*   Report must show each customer along with sum of prices of his orders.
*   It must represent all the information in tabular form.
*   Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt; width: 456pt;"><tbody><tr style="height: 15.75pt;"><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 252.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Customer</span></p></td><td style="border-bottom-color: rgb(102, 102, 102); border-bottom-style: solid; border-bottom-width: 1.5pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; border-top-color: rgb(153, 153, 153); border-top-style: solid; border-top-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 181.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Order Price</span></p></td></tr><tr style="height: 76.5pt;"><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-left-color: rgb(153, 153, 153); border-left-style: solid; border-left-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 4.9pt; padding-right: 4.9pt; vertical-align: middle; width: 252.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&lt;&lt;</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">foreach</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> [in orders</span><br><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&nbsp;&nbsp;&nbsp; .Where(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">c.OrderDate.Year</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;"> == 2015)</span><br><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&nbsp;&nbsp;&nbsp; .</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">GroupBy</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">c.Customer</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">)</span><br><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">&nbsp;&nbsp;&nbsp; .</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">OrderBy</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">(g =&gt; </span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">g.Key.CustomerName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">)]&gt;&gt;&lt;&lt;[</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">Key.CustomerName</span><span style="font-family: Calibri; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="border-bottom-color: rgb(153, 153, 153); border-bottom-style: solid; border-bottom-width: 1pt; border-right-color: rgb(153, 153, 153); border-right-style: solid; border-right-width: 1pt; padding-left: 5.4pt; padding-right: 4.9pt; vertical-align: middle; width: 181.2pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;[Sum(c =&gt; </span><span style="font-family: Calibri; font-size: 11pt;">c.Price</span><span style="font-family: Calibri; font-size: 11pt;">)]&gt;&gt;&lt;&lt;/</span><span style="font-family: Calibri; font-size: 11pt;">foreach</span><span style="font-family: Calibri; font-size: 11pt;">&gt;&gt;</span></p></td></tr></tbody></table>
{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Download In-Table List with Filtering, Grouping, and Ordering Template

Please download the sample In-Table List with Filtering, Grouping, and Ordering document we created in this article:

*   [In-Table List with Filtering, Grouping, and Ordering.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering.xlsx?raw=true) (Template for CustomObject and JSON examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_DB.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_DB.xlsx?raw=true) (Template for DataSet examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_XML.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_XML.xlsx?raw=true) (Template for XML examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists 231a7e3d7d2dbe672478531284592cb6 >}}



#### Database Entities

{{< gist GroupDocsGists ed945fcd5141095bd94cd40c5a097421 >}}



#### Using DataSet

{{< gist GroupDocsGists 556dbf3d5845c5d1b92d1c957b9e4843 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 698082446f03e7b302b9fbdd0e9cd85c >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 559ade9bc8b5860fe564ae4afd4621fe >}}



## In-Table List with Filtering, Grouping, and Ordering in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we 'll not reinvent the wheel to recreate a template for generating an 'In-Table List with Filtering, Grouping, and Ordering' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-table-reports-filtered-ordered-grouped/in-table-list-with-filtering-grouping-and-ordering-in-spreadsheet-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [In-Table List with Filtering, Grouping, and Ordering\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_OpenDocument.ods?raw=true) (use same template for JSON examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_DB\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_DB_OpenDocument.ods?raw=true) (use same template for dataset examples)
*   [In-Table List with Filtering, Grouping, and Ordering\_XML\_OpenDocument.ods](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/In-Table%20List%20with%20Filtering%2C%20Grouping%2C%20and%20Ordering_XML_OpenDocument.ods?raw=true)

### Generating the report

#### Custom Objects

{{< gist GroupDocsGists ad70d05d86c56f437435e7ad46883dc0 >}}



#### Database Entities

{{< gist GroupDocsGists 568db958c0ce91af69f9224773898664 >}}



#### Using DataSet

{{< gist GroupDocsGists 577a989d2169b5ca2d0f7aa137c65530 >}}



#### Using XML DataSource

{{< gist GroupDocsGists f40b94708514c740720f7d96155c7296 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 89f6a1575ad5a8abb78068c14195a720 >}}



### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
