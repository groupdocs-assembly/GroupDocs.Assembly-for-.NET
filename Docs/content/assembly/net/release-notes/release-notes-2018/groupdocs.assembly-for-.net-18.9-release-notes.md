---
id: groupdocs-assembly-for-net-18-9-release-notes
url: assembly/net/groupdocs-assembly-for-net-18-9-release-notes
title: GroupDocs.Assembly for .NET 18.9 Release Notes
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 18.9.{{< /alert >}}

## Major Features

This release of GroupDocs.Assembly brings a new powerful feature: Abilities to change target file format of an assembled document using file extension or explicit specification.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-80  | Add ability to change target file format when assembling Word Processing documents  | Feature  |
| ASSEMBLYNET-81  | Add ability to change target file format when assembling Spreadsheet documents  | Feature  |
| ASSEMBLYNET-82  | Add ability to change target file format when assembling Presentation documents  | Feature  |
| ASSEMBLYNET-83  | Add ability to change target file format when assembling Email documents  | Feature  |
| ASSEMBLYNET-84  | Add ability to change target file format when assembling Text documents  | Feature  |
| ASSEMBLYNET-88  | Add support for Conholdate licenses  | Feature  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 18.9. It includes not only new and obsoleted public methods, but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly which may affect existing code. Any behaviour introduced that could be seen as a regression and modifies existing behaviour is especially important and is documented here.{{< /alert >}}

## Ability to change the target file format of an assembled document 

Starting from the 18.9 version, GroupDocs.Assembly provides abilities to change target file format of an assembled document using file extension or explicit specification. Supported output file formats depending on input file formats can be found at [this page](https://docs.groupdocs.com/display/assemblynet/Introducing+GroupDocs.Assembly+for+.NET#IntroducingGroupDocs.Assemblyfor.NET-SupportedOutputFileFormatsDependingonInputFileFormats).

### Changing target file format using the file extension

**C#**

```csharp
DataSourceInfo dataSourceInfo = new DataSourceInfo(...);
DocumentAssembler assembler = new DocumentAssembler();
 
assembler.AssembleDocument("template.docx", "result.pdf", dataSourceInfo);

```

### Changing target file format using explicit specification

**C#**

```csharp
Stream sourceStream = ...;
Stream targetStream = ...;
 
DataSourceInfo dataSourceInfo = new DataSourceInfo(...);
DocumentAssembler assembler = new DocumentAssembler();
 
assembler.AssembleDocument(sourceStream, targetStream, new LoadSaveOptions(FileFormat.Pdf), dataSourceInfo);


```

### Newly added types

**C#**

```csharp
/// <summary>
/// Provides information on a single data source object to be used to assemble a document from a template.
/// </summary>
public class DataSourceInfo
```

**C#**

```csharp
/// <summary>
/// Specifies the format of a file.
/// </summary>
public enum FileFormat
```

**C#**

```csharp
/// <summary>
/// Specifies additional options for loading and saving of a document to be assembled.
/// </summary>
public class LoadSaveOptions
```

### New members of DocumentAssembler

**C#**

```csharp
/// <summary>
/// Loads a template document from the specified source path, populates the template document with data from
/// the specified single or multiple sources, and stores the result document to the target path using default
/// <see cref="LoadSaveOptions"/>.
/// </summary>
/// <param name="sourcePath">The path to a template document to be populated with data.</param>
/// <param name="targetPath">The path to a result document.</param>
/// <param name="dataSourceInfos">Provides information on data source objects to be used.</param>
public void AssembleDocument(string sourcePath, string targetPath, params DataSourceInfo[] dataSourceInfos)
```

**C#**

```csharp
/// <summary>
/// Loads a template document from the specified source path, populates the template document with data from
/// the specified single or multiple sources, and stores the result document to the target path using the given
/// <see cref="LoadSaveOptions"/>.
/// </summary>
/// <param name="sourcePath">The path to a template document to be populated with data.</param>
/// <param name="targetPath">The path to a result document.</param>
/// <param name="loadSaveOptions">Specifies additional options for document loading and saving.</param>
/// <param name="dataSourceInfos">Provides information on data source objects to be used.</param>
public void AssembleDocument(string sourcePath, string targetPath, LoadSaveOptions loadSaveOptions, params DataSourceInfo[] dataSourceInfos)
```

**C#**

```csharp
/// <summary>
/// Loads a template document from the specified source stream, populates the template document with data from
/// the specified single or multiple sources, and stores the result document to the target stream using default
/// <see cref="LoadSaveOptions"/>.
/// </summary>
/// <param name="sourceStream">The stream to read a template document from.</param>
/// <param name="targetStream">The stream to write a result document.</param>
/// <param name="dataSourceInfos">Provides information on data source objects to be used.</param>
public void AssembleDocument(Stream sourceStream, Stream targetStream, params DataSourceInfo[] dataSourceInfos)
```

**C#**

```csharp
/// <summary>
/// Loads a template document from the specified source stream, populates the template document with data from
/// the specified single or multiple sources, and stores the result document to the target stream using the given
/// <see cref="LoadSaveOptions"/>.
/// </summary>
/// <param name="sourceStream">The stream to read a template document from.</param>
/// <param name="targetStream">The stream to write a result document.</param>
/// <param name="loadSaveOptions">Specifies additional options for document loading and saving.</param>
/// <param name="dataSourceInfos">Provides information on data source objects to be used.</param>
public void AssembleDocument(Stream sourceStream, Stream targetStream, LoadSaveOptions loadSaveOptions, params DataSourceInfo[] dataSourceInfos)
```

### Removed members of DocumentAssembler

**C#**

```csharp
public void AssembleDocument(string sourcePath, string targetPath, object dataSource)
public void AssembleDocument(string sourcePath, string targetPath, object dataSource, string dataSourceName)
public void AssembleDocument(string sourcePath, string targetPath, object[] dataSources, string[] dataSourceNames)
public void AssembleDocument(Stream sourceStream, Stream targetStream, object dataSource)
public void AssembleDocument(Stream sourceStream, Stream targetStream, object dataSource, string dataSourceName)
public void AssembleDocument(Stream sourceStream, Stream targetStream, object[] dataSources, string[] dataSourceNames)
```

### Breaking changes

Removed members of DocumentAssembler were replaced with new ones. Removed members were not marked as obsolete (deprecated) before removal as usual, because this would complicate migration to new members of DocumentAssembler in Java. Thus, we have introduced a breaking change and customers need to migrate their code from removed members of DocumentAssembler to newly added ones when upgrading to the new version of GroupDocs.Assembly.

### Migrating code from removed members of DocumentAssembler to newly added ones

The following table shows how to migrate code from removed members of DocumentAssembler to newly added ones.

| Old Code | New Code |
| --- | --- |
| 
**C#**

```csharp
 assembler.AssembleDocument(sourcePath, targetPath, dataSource);
```





 | 

**C#**

```csharp
assembler.AssembleDocument(sourcePath, targetPath, new DataSourceInfo(dataSource));
```





 |
| 

**C#**

```csharp
 assembler.AssembleDocument(sourcePath, targetPath, dataSource, dataSourceName);
```





 | 

**C#**

```csharp
assembler.AssembleDocument(sourcePath, targetPath, new DataSourceInfo(dataSource, dataSourceName));
```



 |
| 

**C#**

```csharp
object[] dataSources = new object[] { dataSource1, dataSource2 };
string[] dataSourceNames = new string[] { dataSourceName1, dataSourceName2 };
 
assembler.AssembleDocument(sourcePath, targetPath, dataSources, dataSourceNames);
```





 | 

**C#**

```csharp
assembler.AssembleDocument(sourcePath, targetPath, new DataSourceInfo(dataSource1, dataSourceName1),new DataSourceInfo(dataSource2, dataSourceName2));
```



 |
| 

**C#**

```csharp
assembler.AssembleDocument(sourceStream, targetStream, dataSource);
```





 | 

**C#**

```csharp
assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo(dataSource));
```





 |
| 

**C#**

```csharp
assembler.AssembleDocument(sourceStream, targetStream, dataSource, dataSourceName);
```





 | 

**C#**

```csharp
assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo(dataSource, dataSourceName));
```





 |
| 

**C#**

```csharp
object[] dataSources = new object[] { dataSource1, dataSource2 };
string[] dataSourceNames = new string[] { dataSourceName1, dataSourceName2 };
 
assembler.AssembleDocument(sourceStream, targetStream, dataSources, dataSourceNames);
```





 | 

**C#**

```csharp
assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo(dataSource1, dataSourceName1), new DataSourceInfo(dataSource2, dataSourceName2));
```





 |
