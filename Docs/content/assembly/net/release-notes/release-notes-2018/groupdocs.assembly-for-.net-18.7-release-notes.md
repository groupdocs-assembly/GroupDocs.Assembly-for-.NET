---
id: groupdocs-assembly-for-net-18-7-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-7-release-notes
title: GroupDocs.Assembly for .NET 18.7 Release Notes
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.7.{{< /alert >}}

## Major Features

This release adds abilities to insert hyperlinks dynamically for the majority of file formats supported by GroupDocs.Assembly.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-72  | Add ability to insert hyperlinks dynamically for Word Processing documents  | Feature  |
| ASSEMBLYNET-73  | Add ability to insert hyperlinks dynamically for Spreadsheet documents  | Feature  |
| ASSEMBLYNET-74  | Add ability to insert hyperlinks dynamically for Presentation documents  | Feature  |
| ASSEMBLYNET-75  | Add ability to insert hyperlinks dynamically for emails with HTML and RTF bodies  | Feature  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.7. It includes not only new and obsoleted public methods, but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly which may affect existing code. Any behaviour introduced that could be seen as a regression and modifies existing behaviour is especially important and is documented here.{{< /alert >}}

### Inserting Hyperlinks Dynamically 

You can insert hyperlinks to your reports dynamically using link tags. The syntax of a link tag is defined as follows:

```csharp
<<link [uri_expression][display_text_expression]>>

```

Here, **uri\_expression** defines URI for a hyperlink to be inserted dynamically. This expression is mandatory and must return a non-empty value. In turn, **display\_text\_expression** defines text to be displayed for the hyperlink. This expression is optional. If it is omitted or returns an empty value, then during runtime, a value of **uri\_expression** is used as display text as well.
