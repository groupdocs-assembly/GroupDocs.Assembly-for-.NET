---
id: groupdocs-assembly-for-net-19-11-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-11-release-notes
title: GroupDocs.Assembly for .NET 19.11 Release Notes
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.11{{< /alert >}}

## Major Features

Supported dynamic insertion of links to internal document contents and extended the set of supported Markdown features.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-128  | Support dynamic insertion of links to bookmarks for Word Processing documents  | Feature  |
| ASSEMBLYNET-129  | Support dynamic insertion of links to cells for Spreadsheet documents  | Feature  |
| ASSEMBLYNET-130  | Support dynamic insertion of links to slides for Presentation documents  | Feature  |
| ASSEMBLYNET-131  | Support dynamic insertion of links to bookmarks for emails  | Feature  |
| ASSEMBLYNET-133  | Support code blocks and spans for Markdown  | Feature  |
| ASSEMBLYNET-134  | Support strikethrough text for Markdown  | Feature  |
| ASSEMBLYNET-132  | Valid license files are not accepted  | Bug  |

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.11. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

## Public API and Backward Incompatible Changes 

### Extended the set of supported Markdown features

From now on, the following Markdown features are supported when saving assembled Markdown documents to Word Processing formats and saving assembled Word Processing documents and emails to Markdown:

*   [Indented code blocks ](https://spec.commonmark.org/0.29/#indented-code-blocks)
*   [Fenced code blocks](https://spec.commonmark.org/0.29/#fenced-code-blocks)
*   [Inline code spans](https://spec.commonmark.org/0.29/#code-spans)
*   [Strikethrough text](https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet#emphasis)

### Supported dynamic insertion of links to bookmarks for Word Processing documents and emails

You can also insert links to bookmarks to your reports dynamically using link tags. The syntax of a link tag is defined as follows:

```csharp
<<link [uri_or_bookmark_expression][display_text_expression]>>

```

Here, **uri\_or\_bookmark\_expression** defines URI or the name of a bookmark within the same document for a hyperlink to be inserted dynamically. This expression is mandatory and must return a non-empty value.In turn, **display\_text\_expression** defines text to be displayed for the hyperlink. This expression is optional. If it is omitted or returns an empty value, then during runtime, a value of **uri\_or\_bookmark\_expression** is used as display text as well.

{{< alert style="warning" >}}Values of both uri_or_bookmark_expression and display_text_expression can be of any types. During runtime, Object.ToString() is invoked to get textual representations of these expressions’ values, which is useful for expressions of types like Uri, for example. While building a report, uri_or_bookmark_expression and display_text_expression are evaluated and their results are used to construct a hyperlink that replaces the corresponding link tag then. If uri_or_bookmark_expression returns the name of a bookmark in the same document, then the hyperlink navigates to the bookmark. Otherwise, the hyperlink navigates to a corresponding external resource.A link tag cannot be used within a chart.{{< /alert >}}

### Supported dynamic insertion of links to cells for Spreadsheet documents

For Spreadsheet documents, behavior of link tags is changed as follows. If an expression defined within a link tag is evaluated to a cell or cell range reference during runtime, then the tag is replaced with a link to the corresponding cell or cell range.

The following table describes supported formats of cell and cell range references.

| Description | Format | Example |
| --- | --- | --- |
| Reference to a cell within the same worksheet | cell\_name | A1 |
| Reference to a cell in another worksheet | worksheet\_name!cell\_name | Sheet1!A1 |
| Reference to a cell range within the same worksheet | start\_cell\_name:end\_cell\_name | A1:B2 |
| Reference to a cell range in another worksheet | worksheet\_name!start\_cell\_name:end\_cell\_name | Sheet1!A1:B2 |

Following is sample syntax, If the link to cell A1 is required to be inserted:

```csharp
<<link ["A1"] ["Home"]>>
```

### Supported dynamic insertion of links to slides for Presentation documents

For Presentation documents, behavior of link tags is changed as follows. If an expression defined within a link tag is evaluated to a "SlideN" value, where N is a one-based index of a slide within the same Presentation document, then the tag is replaced with a link to the corresponding slide during runtime.

See the example of the syntax as follows:

```csharp
<<link ["Slide1"] ["Home"]>>
```
