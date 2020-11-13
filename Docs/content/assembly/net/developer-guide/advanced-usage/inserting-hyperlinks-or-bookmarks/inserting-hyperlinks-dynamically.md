---
id: inserting-hyperlinks-dynamically
url: assembly/net/inserting-hyperlinks-dynamically
title: Inserting Hyperlinks Dynamically
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.7. or greater{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Inserting Hyperlinks Dynamically

Using GroupDocs.Assembly, you can insert hyperlinks to your reports dynamically using link tags. The syntax of a link tag is defined as follows:

```csharp
<<link [uri_expression][display_text_expression]>>
```

Here, **uri\_expression** defines URI for a hyperlink to be inserted dynamically. This expression is mandatory and must return a non-empty value. In turn, **display\_text\_expression** defines text to be displayed for the hyperlink. This expression is optional. If it is omitted or returns an empty value, then during runtime, a value of **uri\_expression** is used as display text as well. Values of both **uri\_expression** and **display\_text\_expression** can be of any types. During runtime, **Object.ToString()** is invoked to get textual representations of these expressions’ values, which is useful for expressions of types like Uri, for example, while building a report, **uri\_expression** and **display\_text\_expression** are evaluated and their results are used to construct a hyperlink that replaces the corresponding link tag then.

### Dynamic insertion of links to bookmarks 

{{< alert style="info" >}}This feature is supported by version 19.11. or greater{{< /alert >}}

#### Insertion of links for Word Processing documents and Emails

You can also insert links to bookmarks to your reports dynamically using link tags. The syntax of a link tag is defined as follows:

```csharp
<<link [uri_or_bookmark_expression][display_text_expression]>>
```

Here, **uri\_or\_bookmark\_expression** defines URI or the name of a bookmark within the same document for a hyperlink to be inserted dynamically. This expression is mandatory and must return a non-empty value.In turn, **display\_text\_expression** defines text to be displayed for the hyperlink. This expression is optional. If it is omitted or returns an empty value, then during runtime, a value of **uri\_or\_bookmark\_expression** is used as display text as well.

{{< alert style="warning" >}}

* Values of both uri_or_bookmark_expression and display_text_expression can be of any types. During runtime, Object.ToString() is invoked to get textual representations of these expressions’ values, which is useful for expressions of types like Uri, for example. While building a report, uri_or_bookmark_expression and display_text_expression are evaluated and their results are used to construct a hyperlink that replaces the corresponding link tag then. If uri_or_bookmark_expression returns the name of a bookmark in the same document, then the hyperlink navigates to the bookmark. Otherwise, the hyperlink navigates to a corresponding external resource.

* A link tag cannot be used within a chart.

{{< /alert >}}

#### Insertion of links to cells for Spreadsheet documents

For Spreadsheet documents, behavior of link tags is changed as follows. If an expression defined within a link tag is evaluated to a cell or cell range reference during runtime, then the tag is replaced with a link to the corresponding cell or cell range.

The following table describes supported formats of cell and cell range references.

| Description | Format | Example |
| --- | --- | --- |
| Reference to a cell within the same worksheet | cell\_name | A1 |
| Reference to a cell in another worksheet | worksheet\_name!cell\_name | Sheet1!A1 |
| Reference to a cell range within the same worksheet | start\_cell\_name:end\_cell\_name | A1:B2 |
| Reference to a cell range in another worksheet | worksheet\_name!start\_cell\_name:end\_cell\_name | Sheet1!A1:B2 |

Following is sample syntax, If the insertion of the link to cell A1 is required :

```csharp
<<link ["A1"] ["Home"]>>
```

#### Insertion of links to slides for Presentation documents

For Presentation documents, behavior of link tags is changed as follows. If an expression defined within a link tag is evaluated to a "SlideN" value, where N is a one-based index of a slide within the same Presentation document, then the tag is replaced with a link to the corresponding slide during run time.

See the example of the syntax as follows:

```csharp
<<link ["Slide1"] ["Home"]>>
```

## Code Examples and Templates

### Word Processing Documents

{{< gist GroupDocsGists 4087b38f9a93682ffa191599eaacac83 >}}

#### Download

*   [Dynamic Hyperlink.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Dynamic%20Hyperlink.docx)

### Presentation Documents 

{{< gist GroupDocsGists 38e274938c8b7328e250a5d8f1d15092 >}}

#### Download

*   [Dynamic Hyperlink.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Dynamic%20Hyperlink.pptx)

### Spreadsheet Documents 

{{< gist GroupDocsGists 78df21c7cc456a544d46d5b0af95c263 >}}

#### Download

*   [Dynamic Hyperlink.xlsx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Spreadsheet%20Templates/Dynamic%20Hyperlink.xlsx)

### Email Documents

{{< gist GroupDocsGists e5bd8e2bb7406f93ad39dc10eb9f8eb7 >}}

#### Download

*   [Dynamic Hyperlink.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Email%20Templates/Dynamic%20Hyperlink.msg)
