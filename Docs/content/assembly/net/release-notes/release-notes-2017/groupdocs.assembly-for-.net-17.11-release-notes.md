---
id: groupdocs-assembly-for-net-17-11-release-notes
url: assembly/net/groupdocs-assembly-for-net-17-11-release-notes
title: GroupDocs.Assembly for .NET 17.11 Release Notes
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 17.11.{{< /alert >}}

## Major Features

This release of GroupDocs.Assembly comes up with several new features to dynamically manipulate non-textual document elements.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-43  | Support dynamic insertion of images for email messages with HTML body  | Feature |
| ASSEMBLYNET-47 | Add ability to remove selective chart series dynamically  | Feature |
| ASSEMBLYNET-50 | Support dynamic generation of Codablock F barcodes  | Feature |
| ASSEMBLYNET-51 | Support dynamic generation of GS1 Codablock F barcodes  | Feature |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 17.11. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported Dynamic insertion of images for email messages with HTML body

Dynamic insertion of images and barcodes is now supported for MSG, EML, and MHTML documents with HTML body created using Microsoft Outlook or Microsoft Word:

| Feature | MSG, EML, MHTML with HTML body created using Microsoft Outlook or Microsoft Word |
| --- | --- |
| Inserting Images Dynamically | Supported |
| Inserting Barcodes Dynamically | Supported |

### Add ability to remove selective chart series dynamically 

For a chart with dynamic data, you can select which series to include into it dynamically based upon conditions. In particular, this feature is useful when you need to restrict access to sensitive data in chart series for some users of your application. To use the feature, do the following steps:

1.  Declare a chart with dynamic data in the usual way
2.  For series to be removed from the chart based upon conditions dynamically, define the conditions in names of these series using removeif tags having the following syntax  
      
    

```csharp
<<removeif [conditional_expression]>>
```

### Added Support for Codablock F and GS1 Codablock barcodes

The following identifiers can be used to generate Codablock F and GS1 Codablock F barcodes:

| Identifier | Description |
| --- | --- |
| codablockF | Codablock F barcode |
| codablockFGS1 | GS1 Codablock F barcode |
