---
id: groupdocs-assembly-for-net-20-6-release-notes
url: assembly/net/groupdocs-assembly-for-net-20-6-release-notes
title: GroupDocs.Assembly for .NET 20.6 Release Notes
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

Supported working with POT and OTP Presentation document formats and working with ordered lists for Markdown.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-143  | Support ordered lists for Markdown  | Feature  |
| ASSEMBLYNET-154  | Support working with POT and OTP Presentation document formats  | Feature  |
| ASSEMBLYNET-158  | InvalidOperationException is thrown on accessing an empty JSON array  | Bug  |

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 20.6. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

## Public API and Backward Incompatible Changes 

### Supported ordered lists for Markdown

Ordered lists (see chapter 5.3 of [Markdown specification](https://spec.commonmark.org/0.28/)) are supported when saving assembled Markdown documents to Word Processing formats and saving assembled Word Processing documents and emails to Markdown.

### Supported working with POT and OTP Presentation document formats

Support for loading of template POT and OTP Presentation documents and saving of assembled Presentation documents to POT and OTP formats.

#### Loading of a POT (or OTP) template for Presentation document assembly:

```csharp
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument("template.pot", "result.pptx", ...); // For OTP, it might be "template.otp".
```

#### Saving an assembled Presentation document to the POT (or OTP) format using file extension:

```csharp
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument("template.pptx", "result.pot", ...); // For OTP, it might be "result.otp".
```

#### Saving an assembled Presentation document to the POT (or OTP) format using explicit specifying:

```csharp
Stream sourceStream = ...;
Stream targetStream = ...;
 
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(sourceStream, targetStream, new LoadSaveOptions(FileFormat.Pot), ...); // For OTP, FileFormat.Otp should be used.
```
