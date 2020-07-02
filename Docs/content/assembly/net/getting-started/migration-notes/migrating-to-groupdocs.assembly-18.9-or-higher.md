---
id: migrating-to-groupdocs-assembly-18-9-or-higher
url: assembly/net/migrating-to-groupdocs-assembly-18-9-or-higher
title: Migrating to GroupDocs.Assembly 18.9 or Higher
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This document describes the changes to the GroupDocs.Assembly API from version 18.7 to 18.9 or greater that may be of interest to module/application developers. It includes not only new and updated public methods, added & removed classes etc., but also a description of any changes in the behaviour behind the scenes in GroupDocs.Assembly.{{< /alert >}}

### Breaking changes

Removed members of DocumentAssembler were replaced with new ones. Removed members were not marked as obsolete (deprecated) before removal as usual, because this would complicate migration to new members of DocumentAssembler in Java. Thus, we have introduced a breaking change and customers need to migrate their code from removed members of DocumentAssembler to newly added ones when upgrading to the new version of GroupDocs.Assembly.

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

### Migrating code from removed members of DocumentAssembler to newly added ones

The following table shows how to migrate code from removed members of DocumentAssembler to newly added ones.

| Old Code | New Code |
| --- | --- |
| `assembler.AssembleDocument(sourcePath, targetPath, dataSource);` | `assembler.AssembleDocument(sourcePath, targetPath, new DataSourceInfo(dataSource));` |
| `assembler.AssembleDocument(sourcePath, targetPath, dataSource, dataSourceName);` | `assembler.AssembleDocument(sourcePath, targetPath, new DataSourceInfo(dataSource, dataSourceName));` |
| `object[] dataSources = new object[] { dataSource1, dataSource2 };`<br/>`string[] dataSourceNames = new string[] { dataSourceName1, dataSourceName2 };`<br/>`assembler.AssembleDocument(sourcePath, targetPath, dataSources, dataSourceNames);` | `assembler.AssembleDocument(sourcePath, targetPath, new DataSourceInfo(dataSource1, dataSourceName1), new DataSourceInfo(dataSource2, dataSourceName2));` |
| `assembler.AssembleDocument(sourceStream, targetStream, dataSource);` | `assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo(dataSource));` |
| `assembler.AssembleDocument(sourceStream, targetStream, dataSource, dataSourceName);` | `assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo(dataSource, dataSourceName));` |
| `object[] dataSources = new object[] { dataSource1, dataSource2 };`<br/>`string[] dataSourceNames = new string[] { dataSourceName1, dataSourceName2 };`<br/>`assembler.AssembleDocument(sourceStream, targetStream, dataSources, dataSourceNames);` | `assembler.AssembleDocument(sourceStream, targetStream, new DataSourceInfo(dataSource1, dataSourceName1), new DataSourceInfo(dataSource2, dataSourceName2));` |
