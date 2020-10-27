---
id: working-with-outer-document-insertion
url: assembly/net/working-with-outer-document-insertion
title: Working with Outer Document Insertion
weight: 28
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 3.2.0 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

## Inserting Documents Dynamically

You can insert contents of outer documents to your reports dynamically using doc tags. A doc tag denotes a placeholder within a template for a document to be inserted during run-time.

```csharp
<<doc [document_expression]>>
```

{{< alert style="warning" >}}A doc tag can be used almost anywhere in a template document except text-boxes and charts.{{< /alert >}}

An expression declared within a doc tag is used by the engine to load a document to be inserted during run-time. The expression must return a value of one of the following types:

*   A byte array containing document data
*   A Stream instance able to read document data
*   An instance of the Document class
*   A string containing a document URI

{{< alert style="warning" >}}If an expression declared within a doc tag returns a stream object, then the stream is closed by the engine as soon as a corresponding document is loaded.{{< /alert >}}

By default, a document being inserted is not checked against template syntax and is not populated with data. However, you can enable this by using a build switch as follows.

```csharp
<<doc [document_expression] -build>>
```

When a build switch is used, the engine treats a document being inserted as a template that can access the following data available at the scope of a corresponding doc tag:

*   Data sources
*   Variables
*   A contextual object (see "[Using Contextual Object Member Access](https://docs.aspose.com/display/wordsnet/Template+Syntax#nowhere)" for more information)
*   Known external types (see "[Setting up Known External Types](https://docs.aspose.com/display/wordsnet/Template+Syntax#nowhere)"for more information)

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/){{< /alert >}}

### Working with Outer Document Insertion 

#### Download Templates

Get designed and outer templates from here.

*   [OuterDocInsertion.odt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/OuterDocInsertion.odt?raw=true)
*   [OuterDoc.odt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/OuterDocuments/outerDoc.odt?raw=true)

#### The Code

{{< gist GroupDocsGists d86e65933ed173d59f1e5109ff20d33b >}}

### Outer Document Insertion in Word Processing Formats

{{< alert style="info" >}}This feature is supported by version 18.12 or greater.{{< /alert >}}

#### Download Templates

Get designed and outer templates from here.

*   [NestedExternalDocument.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Nested%20External%20Document.docx)
*   [OuterDocument.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/OuterDocuments/OuterDocument.docx)

#### The Code

{{< gist GroupDocsGists db91bbaafa91669b04f305bbf909a9cd >}}

### Outer Document Insertion in Email Formats

{{< alert style="info" >}}This feature is supported by version 18.12 or greater.{{< /alert >}}

#### Download Templates

Get designed and outer templates from here.

*   [NestedExternalDocument.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Email%20Templates/Nested%20External%20Document.msg)
*   [OuterDocument.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/OuterDocuments/OuterDocument.docx)

#### The Code

{{< gist GroupDocsGists 54701f594999260d10a4e6d66fa069e7 >}}
