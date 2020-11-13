---
id: features-overview
url: assembly/net/features-overview
title: Features Overview
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---

GroupDocs.Assembly for .NET is an API that allows you to generate documents in popular *microsoft office* and email file formats based upon template documents and data obtained from various sources including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more. The following is the list of salient features which are offered by the GroupDocs.Assembly API:

*   Multiple data formats support
*   Use formulae & perform sequential data operations
*   Apply upper, lower, capital, firstcap formatting to strings in template syntax
*   Template syntax support formatting of ordinal, cardinal, alphabetic numeric nature
*   Support template documents with custom variables & text comments within template syntax tags
*   Dynamically insert out document content in report
*   Dynamically configure background color of html documents & generate barcode in reports
*   Dynamically insert hyperlinks in reports & apply attributes to email message body
*   Dynamically attach email attachments
*   Support of next field analogue of Microsoft Word
*   Update fields during word processing document assembly
*   Apply formula during spreadsheet assembly
*   Template formatting for numeric, text, image, date-time, chart elements
*   Conditional formatting of template text elements
*   Linq-based template syntax
*   Change format of assembled file using explicit specifications or file extension
*   Auto-removal of empty paragraphs
*   Generate various report types, e.g., charts, image, tables, lists and more
*   Inline template syntax errors in generated documents instead of exception throwing.
*   Load template documents from html with resources and save assembled Word, Excel, PowerPoint & emails to html with resources

Following is the detail given about how the features are offered with their usage and limits.

## Data Sources

### Data Formats

| Feature | Support in GroupDocs.Assembly for .NET |
| --- | --- |
| Database | Supported (using DataSet, DataTable, etc.) |
| XML | Supported (using DataSet.ReadXml or XML deserialization of custom objects) |
| OData | Supported (using deserialization of custom objects through WCF or [third-party .NET libraries](https://en.wikipedia.org/wiki/Open_Data_Protocol#.NET) |
| JSON | Supported (using a third-party component like [Json.NET](http://www.newtonsoft.com/json) on DataSet or custom objects in .NET |
| Custom .NET Objects | Supported |
| Ability to use Spreadsheet as Table of Data | Supported |
| Ability to use Word processing Table as Table of Data | Supported |

### Data Manipulation Capabilities

| Feature | Support in GroupDocs.Assembly for .NET |
| --- | --- |
| Formulas | Supported (using C# syntax subset) |
| Sequential Data Operations (filtering, ordering, grouping, aggregating, etc.) | Supported (through LINQ syntax for data sources of all types) |
| Type Member Invocation | Supported |
| Built-In Data Relation Support | Supported |
| Data Processing Customization | Supported (using external .NET types) |
| External Document Import | Supported |

### Explicit Support for Multiple Data Sources

A new overload [**DocumentAssembler.AssembleDocument**](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassembler/methods/assembledocument/index) is added in the new release. Using this overload user can reference multiple data source objects and their members in the template. Overload accepts an array of data source names and an array of data sources. Learn more about support of multiple data sources in [this](https://docs.groupdocs.com/assembly/net/working-with-multiple-data-sources/) article.

## Template Formats, Elements and Syntax

### Template Document Formats

#### Word Processing Document Formats

Word Processing Document formats including Microsoft Word Document and OpenOffice document formats are supported. For more details, please visit [Support for Word Processing Document Formats]({{< ref "assembly/net/getting-started/supported-document-formats.md#supportfor-word-processing-document-formats" >}}#SupportedDocumentFormats-SupportforWordProcessingDocumentFormats).

#### Spreadsheet Document Formats

Spreadsheet Document formats including Microsoft Excel Workbook and OpenOffice spreadsheet document formats are supported. For more details, please visit [Support for Spreadsheet Document Formats]({{< ref "assembly/net/getting-started/supported-document-formats.md#supportfor-spreadsheet-document-formats" >}}#SupportedDocumentFormats-SupportforSpreadsheetDocumentFormats).

#### Presentation Document Formats

Presentation Document formats including Microsoft PowerPoint Presentation and OpenOffice presentation document formats are supported. For more details, please visit [Support for Presentation Document Formats]({{< ref "assembly/net/getting-started/supported-document-formats.md#supportfor-presentation-document-formats" >}}#SupportedDocumentFormats-SupportforPresentationDocumentFormats).

#### Email Document Formats

Email Document formats is supported. For more details, please visit [Support for Email Document Formats](https://docs.asposeptyltd.com/display/assemblynet/Supported+Document+Formats#SupportedDocumentFormats-EmailDocumentFormats).

#### More Document Formats

HTML and Plain Text document formats are supported by the GroupDocs.Assembly API. More Document formats including PDF, Epub and XPS document formats will be supported in the future releases. For more details, please visit [Support for more Document Formats]({{< ref "assembly/net/getting-started/supported-document-formats.md#supportformore-document-formats" >}}#SupportedDocumentFormats-SupportformoreDocumentFormats).

#### Dynamic Merging of Table Cells

You can merge table cells with equal textual contents within your reports dynamically using **cellMerge** tag. For the detail, please visit [Merging Table Cells Dynamically](https://docs.groupdocs.com/assembly/net/merging-table-cells-dynamically/).

#### Textual Comments within Template Syntax Tags

An optional comment providing a human-readable explanation ignored by the engine

```csharp
<<tag_name [expression] –switch1 –switch2 ... // optional_comment >>
```

#### In-lining of Syntax Error Messages into Templates 

In-lining of syntax error messages into templates instead of exception throwing. For the detail, please visit [Use of In-line Syntax Error Messages into Templates](https://docs.groupdocs.com/assembly/net/use-of-in-line-syntax-error-messages-into-templates/).

### Template Elements

| Feature | Support in GroupDocs.Assembly for .NET |
| --- | --- |
| Formatted Text Blocks | Supported |
| HTML Blocks | Supported |
| Repeated Blocks (including list items and table rows) | Supported |
| Conditional Blocks | Supported (including list items and table rows) |
| Pivot Tables | To be supported |
| Images | Supported |
| Charts | Supported |
| Barcodes (through barcode fonts) | Supported |
| Bound Form Controls | To be supported |
| Hyperlinks to URI or Bookmarks | Supported |

### Template Syntax Formats for Expression Results

#### Specifying String Formats

API provides a useful feature to format strings in template syntax. String formatting such as upper, lower, capital, etc.

| String Format | Description |
| --- | --- |
| lower | Converts a string to lower case ("the string") |
| upper | Converts a string to upper case ("THE STRING") |
| caps | Capitalizes a first letter of every word in a string ("The String") |
| firstCap | Capitalizes the first letter of the first word in a string ("The string") |

#### Specifying Numeric Formats

Using GroupDocs.Assembly for .NET user can also do numeric formatting in template syntax. Numeric formatting such as ordinal, cardinal, alphabetic, etc.

| Number Format | Description |
| --- | --- |
| alphabetic | Formats an integer number as an upper-case letter (A, B, C, ...) |
| roman | Formats an integer number as an upper-case Roman numeral (I, II, III, ...) |
| ordinal | Appends an ordinal suffix to an integer number (1st, 2nd, 3rd, ...) |
| ordinalText | Converts an integer number to its ordinal text representation (First, Second, Third, ...) |
| cardinal | Converts an integer number to its text representation (One, Two, Three, ...) |
| hex | Formats an integer number as hexadecimal (8, 9, A, B, C, D, E, F, 10, 11, ...) |
| arabicDash | Encloses an integer number with dashes (- 1 -, - 2 -, - 3 -, ...) |

#### Variables in Template Documents

API permits users to define variables in template documents as follows:

```csharp
<<var [st = "Hello, "]>><<[st]>><<var[st = "World!"]>><<[st]>>
```

### Support for Outer Document Insertion

You can insert contents of outer documents to your reports dynamically.

### Barcode Image Generation

Barcode images can be generated in reports dynamically.

### Support for Analogue of Microsoft Word NEXT Field

This feature allows user to implement a new <<next>> tag. The tag should be supported only within data bands created by using <<foreach>> tag.

### Setting Background Color Dynamically

For [HTML documents](https://docs.groupdocs.com/assembly/net/bulleted-list-in-html-document/), a background color can not be set dynamically using the backColor tag

```csharp
<<backColor ["red"]>>text with red background<</backColor>>

```

Aside this, you can use the HTML style attribute to achieve this.

### Setting Email Message Attributes Dynamically

In addition to an email message body, you can use GroupDocs.Assembly to set the following email message attributes dynamically using the same template syntax:

*   Subject
*   From
*   To
*   CC
*   BCC

For example, given that recipients is an enumeration of strings "Named Recipient <[named@example.com](mailto:named@example.com)\>" and "[unnamed@example.com](mailto:unnamed@example.com)", you can use the following template to set the To attribute for an email message.

```csharp
<<foreach [r in recipients]>><<[r]>>; <</foreach>>

```

During runtime, the To attribute is set to the following value then.

```csharp
Named Recipient <named@example.com>; unnamed@example.com

```

### Building Email Message Attachments Dynamically

GroupDocs.Assembly enables you to build contents of documents attached to email messages dynamically. To do this, attach template documents (of file formats supported by GroupDocs.Assembly) to your template email message. Then, while assembling the email message, the engine assembles the attached documents as well using the same data sources.

### Ability to Update Fields

API permits users to update fields while assembling Word Processing documents.

### Ability to Calculate Formulas

API allows users to calculate formula while assembling Spreadsheet documents.

### Template Elements Formatting

| Feature | Support in GroupDocs.Assembly for .NET |
| --- | --- |
| Numeric/Date-Time Value Formatting | Supported |
| Text Formatting | Supported |
| Conditional Text Formatting | Supported (only through conditional blocks) |
| Image Formatting | Supported (WYSIWYG) |
| Chart Formatting | Supported (WYSIWYG) |

### Template Syntax

| Feature | Support in GroupDocs.Assembly for .NET |
| --- | --- |
| [LINQ-based](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/) | Supported |
| [Mustache](https://mustache.github.io/mustache.5.html) | To be supported |

### Changing Output File Format 

GroupDocs.Assembly provides a powerful feature to change the target file format of an assembled document using file extension or explicit specification.

### Loading HTML and Saving To HTML with External Resource Files

Load template documents from html with resources and save assembled Word, Excel, PowerPoint & emails to html with resources. Please visit the following articles for the detail:

*   [Loading of template documents from HTML with resources](https://docs.groupdocs.com/assembly/net/loading-of-template-documents-from-html-with-resources/)
*   [Saving to HTML with External Resource Files](https://docs.groupdocs.com/assembly/net/saving-to-html-with-external-resource-files/)

### Numbering Restart in Nested Numbered List 

The GroupDocs.Assembly engine allows restart list numbering within your documents dynamically using *<<restartNum>>* tags. In particular, this feature is useful when working with a nested numbered list within a data band. Please visit [this article](https://docs.groupdocs.com/assembly/net/numbered-list-in-word-processing-document/#numbering-restart-in-nested-numbered-list) for more details

  

  

## Metered Licensing

The GroupDocs.Assembly.Metered public class has been added to provide metered licensing. This feature is supported in GroupDocs.Assembly for .NET 16.12.0.
