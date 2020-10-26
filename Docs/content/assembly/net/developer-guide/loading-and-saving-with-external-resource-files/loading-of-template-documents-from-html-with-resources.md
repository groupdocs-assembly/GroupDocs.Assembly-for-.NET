---
id: loading-of-template-documents-from-html-with-resources
url: assembly/net/loading-of-template-documents-from-html-with-resources
title: Loading of template documents from HTML with resources
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 19.5. or greater{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

GroupDocs.Assembly provides a great feature which is loading of HTML template documents referencing external resource files to be assembled and saved to non-HTML formats. 

Since version 19.5, when loading an HTML document from a file, its containing folder is used as a base URI to resolve external resource files' relative URIs to absolute ones by default as shown in the following example. The following code snippet shows the way of API usage:

{{< gist GroupDocsGists bfd1b9daaa9a1b0a020a20a3f56413f2 >}}

The following code snippet shows the default folder used to resolve external resource files' relative URIs are overridden by an explicitly specified folder.

{{< gist GroupDocsGists bd0d6a7a7cca104bbe73c4938f2a6cf7 >}}


