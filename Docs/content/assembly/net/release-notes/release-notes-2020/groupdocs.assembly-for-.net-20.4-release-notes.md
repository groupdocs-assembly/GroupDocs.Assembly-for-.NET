---
id: groupdocs-assembly-for-net-20-4-release-notes
url: assembly/net/groupdocs-assembly-for-net-20-4-release-notes
title: GroupDocs.Assembly for .NET 20.4 Release Notes
weight: 80
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

Provided an option to fit an image within textbox bounds while maintaining ratio.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-149  | Provide an option to fit an image within textbox bounds while maintaining ratio for Word Processing documents  | Feature  |
| ASSEMBLYNET-150  | Provide an option to fit an image within textbox bounds while maintaining ratio for emails with RTF bodies  | Feature  |
| ASSEMBLYNET-151  | Provide an option to fit an image within textbox bounds while maintaining ratio for Spreadsheet documents  | Feature  |
| ASSEMBLYNET-152  | Provide an option to fit an image within textbox bounds while maintaining ratio for Presentation documents  | Feature  |

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 20.4. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

## Public API and Backward Incompatible Changes 

### Provided an option to fit an image within textbox bounds while maintaining ratio

To keep the size of the textbox and stretch the image within bounds of the textbox preserving the ratio of the image, use the *keepRatio* switch as follows.

```csharp
<<image [image_expression] -keepRatio>>
```
