---
id: groupdocs-assembly-for-net-19-5-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-5-release-notes
title: GroupDocs.Assembly for .NET 19.5 Release Notes
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.5.{{< /alert >}}

## Major Features

Supported loading of template documents from HTML with resources and saving of assembled documents to HTML with resources.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-86  | Support saving of assembled Word Processing documents to HTML with resources  | Feature  |
| ASSEMBLYNET-87  | Support loading of template documents from HTML with resources  | Feature  |
| ASSEMBLYNET-109  | Support saving of assembled Spreadsheet documents to HTML with resources  | Feature  |
| ASSEMBLYNET-110  | Support saving of assembled Presentation documents to HTML with resources  | Feature  |
| ASSEMBLYNET-111  | Support saving of assembled emails to HTML with resources  | Feature  |
| ASSEMBLYNET-114  | An exception is thrown on saving a document to DOC or PDF assembled from a DOCX template containing charts  | Bug  |

## Public API and Backward Incompatible Changes 

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.5. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported loading of template documents from HTML with resources

From now, the GroupDocs.Assembly supports loading of HTML template documents referencing external resource files to be assembled and saved to non-HTML formats.

#### Changed behavior of DocumentAssembler.AssembleDocument overloads accepting file paths

From now, when loading an HTML document from a file, its containing folder is used as a base URI to resolve external resource files' relative URIs to absolute ones by default as shown in the following example.

```csharp
string myDir = ...
DocumentAssembler assembler = new DocumentAssembler();
 
assembler.AssembleDocument(myDir + "TestWordsResourceLoad.htm", myDir + "TestWordsResourceLoad_Out.docx", new DataSourceInfo("It should be a jeep image.", "value"));
```

After running the code, the TestWordsResourceLoad\_Out.docx contains an image, which relative URI is resolved from myDir.

#### Added a new property to LoadSaveOptions

The following property was added.

```csharp
/// <summary>
/// Gets or sets a base URI to resolve external resource files' relative URIs to absolute ones while loading an HTML 
/// template document to be assembled and saved to a non-HTML format. The default value is an empty string.
/// </summary>
/// <remarks>
/// <para>
/// When loading an HTML document from a file, its containing folder is used as a base URI by default, which cannot 
/// happen when loading an HTML document from a stream. Set this property to specify a base URI when loading an HTML 
/// document from a stream or to override the default base URI when loading an HTML document from a file.
/// </para>
/// <para>
/// A value of this property is ignored in the following cases:
/// </para>
/// <list type="bullet">
/// <item>
/// An HTML document being loaded contains a BASE HTML element providing a base URI.
/// </item>
/// <item>
/// An HTML document being loaded is to be assembled and saved to HTML (external resource files are not loaded and
/// relative URIs are not changed then).
/// </item>
/// </list>
/// </remarks>
public string ResourceLoadBaseUri { get; set; }
```

In the following example, the default folder used to resolve external resource files' relative URIs is overridden by an explicitly specified folder.

```csharp
string myDir = ...
DocumentAssembler assembler = new DocumentAssembler();
 
LoadSaveOptions loadSaveOptions = new LoadSaveOptions();
loadSaveOptions.ResourceLoadBaseUri = myDir + "Alternative";
 
assembler.AssembleDocument(myDir + "TestWordsResourceLoad.htm", myDir + "TestWordsResourceLoad_Out.docx", loadSaveOptions, new DataSourceInfo("It should be a sport car image.", "value"));
```

After running the code, the TestWordsResourceLoad\_Out.docx contains an image, which relative URI is resolved from the specified alternative folder.

{{< alert style="warning" >}}Download ResourceLoad.zip to get the template documents.{{< /alert >}}

### Supported saving of assembled documents to HTML with resources

From now, the GroupDocs.Assembly supports saving of external resource files while an assembled document loaded from a non-HTML format is being saved to HTML.

#### Changed behavior of DocumentAssembler.AssembleDocument overloads accepting file paths

From now, when saving an assembled document to an HTML file, by default, external resource files are stored to a folder having the same name as the HTML file without extension plus the "\_files" suffix. This folder is located in the same folder as the HTML file. The following example illustrates the scenario.

```csharp
string myDir = ...
DocumentAssembler assembler = new DocumentAssembler();

assembler.AssembleDocument(myDir + "TestWordsResourceSave.docx", myDir + "TestWordsResourceSave_Out.htm", new DataSourceInfo("Hello!", "value"));
```

After running the code, the folder "TestWordsResourceSave\_Out\_files" is created in myDir and HTML resource files (such as images and stylesheets) are stored in this folder.

#### Added a new property to LoadSaveOptions

The following property was added.

```csharp
/// <summary>
/// Gets or sets a path to a folder to store external resource files while an assembled document loaded from a non-HTML 
/// format is being saved to HTML. The default value is an empty string.
/// </summary>
/// <remarks>
/// <para>
/// By default, when saving an assembled document to an HTML file, external resource files are stored to a folder
/// having the same name as the HTML file without extension plus the "_files" suffix. This folder is located in
/// the same folder as the HTML file. However, this cannot be done when saving an assembled document to an HTML stream.
/// Set this property to specify a path to a folder to store external resource files when saving an assembled document
/// to an HTML stream or to override the default folder when saving an assembled document to an HTML file.
/// </para>
/// <para>
/// A value of this property is ignored if an assembled document being saved to HTML was loaded from HTML as well
/// (external resource files are not stored and links to them are not changed then).
/// </para>
/// </remarks>
public string ResourceSaveFolder { get; set; }
```

In the following example, the default folder used to store HTML resource files is overridden by an explicitly specified folder.

```csharp
string myDir = ...
DocumentAssembler assembler = new DocumentAssembler();

LoadSaveOptions loadSaveOptions = new LoadSaveOptions();
loadSaveOptions.ResourceSaveFolder = myDir + "Alternative";

assembler.AssembleDocument(myDir + "TestWordsResourceSave.docx", myDir + "TestWordsResourceSave_Out.htm", loadSaveOptions, new DataSourceInfo("Hello!", "value"));
```

After running the code, the resource files (such as images and stylesheets) for the saved HTML document are stored to the specified alternative folder.

{{< alert style="warning" >}}Download ResourceSave.zip to get the template documents.{{< /alert >}}
