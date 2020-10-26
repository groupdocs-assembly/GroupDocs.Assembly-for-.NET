---
id: saving-to-html-with-external-resource-files
url: assembly/net/saving-to-html-with-external-resource-files
title: Saving to HTML with External Resource Files
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 19.5. or greater{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

Since version 19.5, The GroupDocs.Assembly provides a great feature which is saving of external resource files while an assembled document loaded from a non-HTML format is being saved to HTML.

From now on, when saving an assembled document to an HTML file, by default, external resource files are stored to a folder having the same name as the HTML file without extension plus the "\_files" suffix. This folder is located in the same folder as the HTML file. The following example illustrates the scenario. The following code snippet shows the way of API usage:

{{< gist GroupDocsGists 63210a271dd6e3f960b86637e15e3d5a >}}



In the following example, the default folder used to store HTML resource files is overridden by an explicitly specified folder.

{{< gist GroupDocsGists b530fa88fa2e578219aa33f7e064c49d >}}


