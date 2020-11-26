---
id: groupdocs-assembly-for-net-20-11-release-notes
url: assembly/net/groupdocs-assembly-for-net-20-11-release-notes
title: GroupDocs.Assembly for .NET 20.11 Release Notes
weight: 40
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

GroupDocs.Assembly for .NET 20.11 comes up with support of .NET Standard 2.0 as well as several improvements and bug fixes in working with charts, Markdown, JSON data sources, and other areas.

## Full List of Features Covering all Changes in this Release

| Key             | Summary                                                      | Category |
| :-------------- | :----------------------------------------------------------- | :------- |
| ASSEMBLYNET-153 | Support for .NET Core 3.0                                    | Feature  |
| ASSEMBLYNET-172 | Support inline links for Markdown                            | Feature  |
| ASSEMBLYNET-173 | Support inline images for Markdown                           | Feature  |
| ASSEMBLYNET-178 | API support for .NET Core 5.0                                | Feature  |
| ASSEMBLYNET-165 | Nullable int is not recognized for JsonDataSource            | Bug      |
| ASSEMBLYNET-166 | Chart legend entries removed in a template appear in a result Word Processing document | Bug      |
| ASSEMBLYNET-167 | Tag "removeif" throws NullReferenceException for a Word Processing document | Bug      |
| ASSEMBLYNET-168 | JsonDataSource cannot access a property having the same name as its containing object | Bug      |
| ASSEMBLYNET-169 | Tag "cellMerge" with empty cells throws ArgumentOutOfRangeException for a Word Processing document | Bug      |
| ASSEMBLYNET-170 | Tag "backColor" adds unwanted carriage return for a Word Processing document | Bug      |
| ASSEMBLYNET-171 | JsonDataSource throws DuplicateNameException                 | Bug      |

## Public API and Backward Incompatible Changes 

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 20.9. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported Inline Links and Inline Images for Markdown

From now on, inline links and inline images (see [Inline Link](https://spec.commonmark.org/0.29/#inline-link)) are supported when saving assembled Markdown documents to Word Processing formats and saving assembled Word Processing documents and emails to Markdown.

### Supported Working with .NET Standard 2.0 Implementations

Starting with version 20.11, GroupDocs.Assembly supports working with .NET Standard 2.0 implementations such as .NET Core 2.0 and later. For more information on .NET Standard 2.0 implementations, see [.NET Standard documentation](https://docs.microsoft.com/en-us/dotnet/standard/net-standard#net-implementation-support).

#### Working with GroupDocs.Assembly for .NET Standard 2.0 in Your Project

GroupDocs.Assembly is built on top of several Aspose products. Some of the products use [SkiaSharp](https://www.nuget.org/packages/SkiaSharp/) to deal with graphics-related tasks within .NET Standard 2.0, whereas the others prefer [System.Drawing.Common](https://www.nuget.org/packages/System.Drawing.Common/). That is why, GroupDocs.Assembly depends on the both packages.

A recommended way to deal with these and other GroupDocs.Assembly for .NET Standard 2.0 dependencies is installing of [GroupDocs.Assembly NuGet package](https://www.nuget.org/packages/GroupDocs.Assembly/). Otherwise, you can install all dependencies listed at [the NuGet package page](https://www.nuget.org/packages/GroupDocs.Assembly/) manually.

Depending on OS you work on, you may need to install some additional dependencies. Learn more about these dependencies in the article ["How to Work with GroupDocs.Assembly for .NET Standard 2.0"](https://docs.groupdocs.com/assembly/net/how-to-work-with-groupdocs.assembly-for-.net-standard-2.0/).
