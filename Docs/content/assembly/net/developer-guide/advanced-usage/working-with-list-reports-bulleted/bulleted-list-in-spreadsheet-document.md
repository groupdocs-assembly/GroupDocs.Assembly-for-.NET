---
id: bulleted-list-in-spreadsheet-document
url: assembly/net/bulleted-list-in-spreadsheet-document
title: Bulleted List in Spreadsheet Document
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Bulleted List report in Spreadsheet Document based on the use case: [Working with a Business Case](https://docs.groupdocs.com/assembly/net/working-with-a-business-case/).{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Bulleted List in Microsoft Excel Document

### Creating a Bulleted List

Practising the following steps you can insert Bulleted List in MS Excel 2013.  
Adding bulleted list in Microsoft Excel is different than Microsoft Word. Moreover, there are two ways to apply bulleted list in Microsoft Excel:

*   In-Cell List
*   Multiple-Cell List

Apply the following steps to build the In-Cell template

1.  Add a new Workbook.
2.  Press Insert Tab (at the top-bar).
3.  Add bullet symbol by clicking on Symbol icon.
4.  Save the Document.

### Reporting Requirement

As a report developer, you are required to share the information of the products with the following key requirements:

*   A report must show all the products in a bulleted list format.
*   A Report must be generated in the Spreadsheet Document.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

**In-Cell List**

```csharp
We provide support for the following products:
<<foreach [in products]>>
-          <<[ProductName]>><</foreach>>
```

**Multiple-Cell List**

```csharp
We provide support for the following products:
<<foreach [in products]>>-          <<[ProductName]>>
```

close the foreach loop in next column

```csharp
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Bulleted List Template

Please download the sample Bulleted List document we created in this article:

*   [Bulleted List.xlsx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Bulleted%20List.xlsx?raw=true) (Template for DataBase, DataSet and JSON examples)
*   [Bulleted List\_XML.xlsx](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Bulleted%20List_XML.xlsx?raw=true) (Template for XML Data examples)

### Generating The Report

#### Custom Objects

{{< gist GroupDocsGists ec5f7ab4cb1741cc2e33b14de6ec576c >}}



#### Database Entities

{{< gist GroupDocsGists e2848d87be2cbdb9152ff5e47d5f96a4 >}}



#### Using DataSet

{{< gist GroupDocsGists 2ee710544a43ce46cffac4a13bb80d62 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 7f5d0b98132d4e3275e3546fac0b0b41 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 376d39743cf0d06b2288644fa199cf48 >}}



## Bulleted List in OpenOffice Spreadsheet Document

### Creating the Template

OpenDocument Spreadsheet (ODS) is a spreadsheet document format which can be used as an alternative to Microsoft Excel Document (XLS/XLSX) formats. Since ODS is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODS, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating a 'Bulleted List' report in ODS format. Instead, we'll save the existing template to ODS format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in previous topic.
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Spreadsheet" from "Save As Type" dropdown.
4.  Click "Save".

### Download Template

*   [Bulleted List\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Bulleted%20List_OpenDocument.ods?raw=true) (use the same template for database, dataset and JSON examples)
*   [Bulleted List\_XML\_OpenDocument.ods](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Bulleted%20List_XML_OpenDocument.ods?raw=true)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists cde401351d5b1ee82fdd4773703c7423 >}}



#### Database Entities

{{< gist GroupDocsGists 2766a6ab2316cb857403835d643aa655 >}}



#### Using DataSet

{{< gist GroupDocsGists 94b09f9afdac5ca6b81bd4c95a8d7f3b >}}



#### Using XML DataSource

{{< gist GroupDocsGists a5cbb073ae455ab99f66038dc7a7fa58 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 34a97e1b0ea144dda91e167ccbfc2da8 >}}



### ODS Template and Report in Apache OpenOffice

In order to check compatibility of ODS between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODS template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODS report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
