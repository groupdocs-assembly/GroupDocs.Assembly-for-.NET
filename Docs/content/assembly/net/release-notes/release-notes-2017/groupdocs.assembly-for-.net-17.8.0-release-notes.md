---
id: groupdocs-assembly-for-net-17-8-0-release-notes
url: assembly/net/groupdocs-assembly-for-net-17-8-0-release-notes
title: GroupDocs.Assembly for .NET 17.8.0 Release Notes
weight: 4
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 17.8.0{{< /alert >}}

## Major Features

Document assembly for email file formats is supported.

## Full List of Issues Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-32 | Support document assembly for MHTML file formats | Feature |
| ASSEMBLYNET-33 | Support document assembly for MSG file formats | Feature |
| ASSEMBLYNET-34 | Support document assembly for EML file formats | Feature |
| ASSEMBLYNET-35 | Support document assembly for EMLX file formats | Feature |
| ASSEMBLYNET-40 | Support building of charts located in chart worksheets | Feature |
| ASSEMBLYNET-41 | Support building of charts with series names located in referenced cells | Feature |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 17.8.0. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported Common Features for Email File Formats

All common textual and tabular template features such as expressions, variables, data bands, conditional blocks, and others are supported to the full extent.

### Partially Supported or Not Supported Features for Email File Formats

The following table provides information on those GroupDocs.Assembly features that are partially supported or not supported at all for email file formats.

| Feature | MSG with RTF Body | MSG, EML, MHTML with HTML Body Created Using Microsoft Outlook or Microsoft Word | MSG, EML, EMLX, MHTML with HTML Body Created Outside Microsoft Outlook or Microsoft Word | MSG, EML, EMLX, MHTML with Plain Text Body |
| --- | --- | --- | --- | --- |
| Building Charts Dynamically | Supported (outside loops and conditions) | Supported (outside loops and conditions) | Not supported | Not supported |
| Inserting Documents Dynamically | Supported | Supported | To be supported | Not supported |
| Inserting Images Dynamically | To be supported | To be supported | Supported | Not supported |
| Inserting Barcodes Dynamically | To be supported | To be supported | To be supported | Not supported |

### Setting Email Message Attributes Dynamically

In addition to an email message body, you can use GroupDocs.Assembly to set the following email message attributes dynamically using the same template syntax:

*   Subject
*   From
*   To
*   CC
*   BCC

For example, given that `recipients` is an enumeration of strings `"Named Recipient <named@example.com>"` and `"unnamed@example.com"`, you can use the following template to set the `To` attribute for an email message.

```csharp
<<foreach [r in recipients]>><<[r]>>; <</foreach>>

```

During runtime, the `To` attribute is set to the following value then.

```csharp
Named Recipient <named@example.com>; unnamed@example.com

```

### Building Email Message Attachments Dynamically

GroupDocs.Assembly enables you to build contents of documents attached to email messages dynamically. To do this, attach template documents (of file formats supported by GroupDocs.Assembly) to your template email message. Then, while assembling the email message, the engine assembles the attached documents as well using the same data sources.
