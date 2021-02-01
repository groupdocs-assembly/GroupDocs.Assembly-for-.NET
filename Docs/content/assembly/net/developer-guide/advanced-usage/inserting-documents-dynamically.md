---
id: inserting-documents-dynamically
url: assembly/net/inserting-documents-dynamically
title: Inserting Documents Dynamically
weight: 43
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 20.3 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

{{< alert style="warning" >}}Dynamic insertion of documents from Base64-encoded bytes is available for file formats where dynamic document insertion is available for Word Processing documents and emails with HTML and RTF bodies only.{{< /alert >}}

You can insert contents of outer documents to your reports dynamically using `doc` tags. A `doc` tag denotes a placeholder within a template for a document to be inserted during runtime.

Syntax of a `doc` tag is defined as follows:

```csharp
<<doc [document_expression]>>
```

And simply call the assembler method to generate report like following code snippets:

{{< gist GroupDocsGists b06d6c5655c84fc772c6411d66016943 InsertDocumentDynamicallyInWord_20.3.cs >}}

{{< alert style="warning" >}}A `doc` tag can be used almost anywhere in a template document except textboxes and charts.{{< /alert >}}

An expression declared within a `doc` tag is used by the assembler to load a document to be inserted during runtime. The expression must return a value of one of the following types:

*   A byte array containing document data
*   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read document data
*   An instance of the Document class
*   A string containing a document URI, path, or Base64-encoded document data

While building a report, an expression declared within a `doc` tag is evaluated and its result is used to load a document which content replaces the `doc` tag then.

{{< alert style="warning" >}}If an expression declared within a `doc` tag returns a stream object, then the stream is closed by the assembler as soon as a corresponding document is loaded.{{< /alert >}}

By default, while inserting content of an outer document, the engine applies corresponding styles of a template document. This makes content of a result document look more consistent. However, you can keep source formatting for content being inserted by using a `sourceStyles` switch as shown in the following snippet:

```csharp
<<doc [document_expression] -sourceStyles>>
```

By default, a document being inserted is not checked against template syntax and is not populated with data. However, you can enable this by using a `build` switch as follows:

```csharp
<<doc [document_expression] -build>>
```

When a `build` switch is used, the assembler treats a document being inserted as a template that can access the following data available at the scope of a corresponding `doc` tag:

*   Data sources
*   Variables
*   A contextual object (see [“Using Contextual Object Member Access”](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-contextual-object-member-access) for more information)
*   Known external types (see [“Setting up Known External Types”](https://docs.groupdocs.com/assembly/net/groupdocs-assembly-engine-apis/#setting-up-known-external-types) for more information)
