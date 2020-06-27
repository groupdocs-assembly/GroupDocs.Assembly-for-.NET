---
id: groupdocs-assembly-for-net-18-12-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-12-release-notes
title: GroupDocs.Assembly for .NET 18.12 Release Notes
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.12.{{< /alert >}}

## Major Features

Supported document assembly of external documents being dynamically inserted for Word Processing and Email formats.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-94 | Support document assembly of external documents being dynamically inserted for Word Processing formats  | Feature  |
| ASSEMBLYNET-95  | Support document assembly of external documents being dynamically inserted for Email formats  | Feature  |
| ASSEMBLYNET-96  | An evaluation mark is added to a nested document being inserted dynamically  | Bug  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.12. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Support document assembly of external documents being dynamically inserted for Word Processing and Email Formats

You can insert contents of outer documents to your reports dynamically using doc tags. A doc tag denotes a placeholder within a template for a document to be inserted during runtime. The syntax of a doc tag is defined as follows.

```csharp
<<doc [document_expression]>>

```

{{< alert style="warning" >}}A doc tag can be used almost anywhere in a template document except textboxes and charts.{{< /alert >}}

An expression declared within a doc tag is used by the engine to load a document to be inserted during runtime. The expression must return a value of one of the following types:

*   A byte array containing document data
*   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read document data
*   A string containing a document URI

While building a report, an expression declared within a doc tag is evaluated and its result is used to load a document which content replaces the doc tag then.

{{< alert style="warning" >}}If an expression declared within a doc tag returns a stream object, then the stream is closed by the engine as soon as a corresponding document is loaded.{{< /alert >}}

By default, a document being inserted is not checked against template syntax and is not populated with data. However, you can enable this by using a build switch as follows.

```csharp
<<doc [document_expression] -build>>

```

When a build switch is used, the engine treats a document being inserted as a template that can access the following data available at the scope of a corresponding doc tag:

*   Data sources
*   Variables
*   A contextual object 
*   Known external types
