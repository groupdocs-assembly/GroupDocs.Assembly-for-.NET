---
id: inserting-bookmarks-dynamically
url: assembly/net/inserting-bookmarks-dynamically
title: Inserting Bookmarks Dynamically
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 20.1 or greater{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Inserting Hyperlinks Dynamically

### Dynamic insertion of bookmarks for Word Processing documents and emails with HTML and RTF bodies

You can insert bookmarks to your reports dynamically using bookmark tags. Syntax of a bookmark tag is defined as follows.

```csharp
<<bookmark [bookmark_expression]>>
bookmarked_content
<</bookmark>>
```

Here, **bookmark\_expression** defines the name of a bookmark to be inserted during run-time. This expression is mandatory and must return a non-empty value. While building a report, **bookmark\_expression** is evaluated and its result is used to construct a bookmark start and end that replace corresponding opening and closing bookmark tags respectively.

{{< alert style="warning" >}}A bookmark tag cannot be used within a chart.{{< /alert >}}

The following code snippet shows how simple you call the document assembler to generate report from the template:

{{< gist GroupDocsGists fe5688b6d99e17223084d5098e9324fa DynamicBookmarkInsertionWord.cs >}}



### Dynamic naming of cell ranges for Spreadsheet documents

Template syntax and usage of Document Assembler is same as above section for dynamic naming of cell ranges for Spreadsheet documents. The difference is that it results to insertion of named cell ranges instead of bookmarks.

#### Download Templates

*   [Dynamic Bookmarks.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Dynamic%20Hyperlink.docx)
*   [Dynamic Cell Ranges.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Dynamic%20Hyperlink.docx)
 