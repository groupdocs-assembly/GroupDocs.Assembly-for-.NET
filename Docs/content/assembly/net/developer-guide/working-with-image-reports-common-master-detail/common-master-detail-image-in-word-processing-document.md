---
id: common-master-detail-image-in-word-processing-document
url: assembly/net/common-master-detail-image-in-word-processing-document
title: Common Master-Detail Image in Word Processing Document
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Common Master-Detail Image report in Word Processing Document format based on the use case: Working with a Business Case{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Common Master-Detail Image in Microsoft Word Document

### Creating a Common Master-Detail Image

Please follow below steps you can create Common Master-Detail Template in MS Word 2013.

1.  Insert the desired shape to display image in it.
2.  Go to Insert Tab and select shape by clicking on Shape Icon.
3.  Save your Document.

### Reporting Requirement

As a report developer, you are required to represent the information of the customers and products with the following key requirements:

*   Report must show customers' picture and name.
*   It must associate the customers with their products.
*   Report must be generated in the Word Processing Document.  
      
    

{{< alert style="info" >}}See how to use images in MS Word here.{{< /alert >}}

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

<<foreach \[in customers\]>>

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; float: left; margin-top: 0pt; margin-right: 9pt; margin-bottom: 0pt; margin-left: 9pt; width: 132.15pt;"><tbody><tr style="height: 112pt;"><td style="border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 120.6pt;"><p style="font-size: 11pt; line-height: 108%; margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt;"><span style="font-family: Calibri; font-size: 11pt;">&lt;&lt;image [</span><span style="font-family: Calibri; font-size: 11pt;">customer.</span><span style="font-family: Calibri; font-size: 11pt;">Photo</span><span style="font-family: Calibri; font-size: 11pt;">]&gt;&gt;</span></p></td></tr></tbody></table>

<<\[CustomerName\]>>

Products: <<foreach \[in Order\]>><<\[IndexOf() != 0 ? ", " : ""\]>><<\[Product.ProductName\]>><</foreach>>

<</foreach>>

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: Working with GroupDocs.Assembly Engine{{< /alert >}}

### Download Common Master-DetailTemplate

Please download the sample Common Master-Detail document we created in this article:

*   [Common Master-Detail.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail.docx?raw=true) (Template for CustomObject and JSON examples)
*   [Common Master-Detail\_DB.docx](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail_DB.docx?raw=true) (Template for DataSet, DataBase and XML examples)

### Generating The Report

#### Custom Objects

//<!\[CDATA\[ function debug() { } // \]\]>//<!\[CDATA\[ // \]\]>//<!\[CDATA\[ // \]\]>//<!\[CDATA\[ // \]\]>//<!\[CDATA\[ // \]\]>//<!\[CDATA\[ Cloak.closeHTML = "<img src=\\'/download/resources/net.customware.confluence.plugin.composition:toggle-cloak/img/navigate\_down\_10.gif\\'/>"; Cloak.openHTML = "<img src=\\'/download/resources/net.customware.confluence.plugin.composition:toggle-cloak/img/navigate\_right\_10.gif\\'/>"; Cloak.toggleZone = true; Cloak.memoryDuration = 0; Cloak.memoryPrefix = "contentId:34439204"; Cloak.memoryPath = "/"; // \]\]>.cloakToggle { cursor: pointer; }//<!\[CDATA\[ // \]\]>//<!\[CDATA\[ // \]\]>//<!\[CDATA\[ Deck.memoryDuration = 0; Deck.memoryPrefix = "contentId:34439204"; Deck.memoryPath = "/"; // \]\]>{{< gist GroupDocsGists f659589cae0f24c5500f9e2fd39c8ff5 >}}



#### Database Entities

{{< gist GroupDocsGists 0f0827a650b0b0e4db55ddca1efa8033 >}}



#### Using DataSet

{{< gist GroupDocsGists 004ac14e5a96b877eefe0b42a7e6e5ef >}}



#### Using XML DataSource

{{< gist GroupDocsGists a8303e24fefbd71ceff390a32e13668b >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 721f9f83f25c880bc9f57383f17d0d6e >}}



## Common Master-Detail in OpenOffice Document

### Creating the Template

OpenDocument Text (ODT) is a word processing document format which can be used as an alternative to Microsoft Word Document (DOC/DOCX) formats. Since ODT is not a Microsoft Proprietary format, there are multiple software (including Microsoft Office and Apache OpenOffice) available to create, open, edit and save this format. For more information on the available software to work with ODT, please visit [wikipedia article](https://en.wikipedia.org/wiki/OpenDocument#Software).

In this topic, we'll not reinvent the wheel to recreate a template for generating an 'Common Master-Detail' report in ODT format. Instead, we'll save the existing template to ODT format using Microsoft Office. In order to achieve this; assuming you are using Microsoft Office 2010, please follow below steps:

1.  Open existing template we created in [previous topic]({{< ref "assembly/net/developer-guide/working-with-image-reports-common-master-detail/common-master-detail-image-in-word-processing-document.md" >}})
2.  Click "File" and select "Save As".
3.  Select "OpenDocument Text" from "Save As Type" drop down.
4.  Click "Save".

### Download Template

*   [Common Master-Detail\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail_OpenDocument.odt?raw=true) (use same template for JSON examples)
*   [Common Master-Detail\_DB\_OpenDocument.odt](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Common%20Master-Detail_DB_OpenDocument.odt?raw=true) (use same template for database and dataset and XML examples)

### Generating the Report

#### Custom Objects

{{< gist GroupDocsGists 9e3a02f4a1f5b7ef54ccb059dd96a039 >}}



#### Database Entities

{{< gist GroupDocsGists b271cec9679d51438f4990085f37ba02 >}}



#### Using DataSet

{{< gist GroupDocsGists 3d522b9b2cae2e11647d4089a4f2dfd7 >}}



#### Using XML DataSource

{{< gist GroupDocsGists 4a676c7c9c231e5fc4f489d36aa8d370 >}}



#### Using JSON DataSource

{{< gist GroupDocsGists 352b28ee4af2e26ca54a188d5bdbfd16 >}}



### ODT Template and Report in Apache OpenOffice

In order to check compatibility of ODT between Microsoft Office 2010 and Apache OpenOffice 4.1.2, we performed below tests:

*   We opened the ODT template created through Microsoft Office 2010 in Apache OpenOffice 4.1.2. The template opened successfully in Apache OpenOffice without any issues or formatting losses.
*   We opened the ODT report generated through GroupDocs.Assembly in Apache OpenOffice 4.1.2. The report opened successfully in Apache OpenOffice without any issues or formatting losses.
