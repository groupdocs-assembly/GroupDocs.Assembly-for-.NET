---
id: migrating-to-groupdocs-assembly-3-0-0-or-higher
url: assembly/net/migrating-to-groupdocs-assembly-3-0-0-or-higher
title: Migrating to GroupDocs.Assembly 3.0.0 or Higher
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Reasons to Migrate

Just to endeavor the report generation task more smoothly, we released **GroupDocs.Assembly for .NET 3.0.0**.

### Namespace(s)

In **GroupDocs.Assembly for .NET 3.0.0** only a single **using GroupDocs.Assembly;** namespace is required to generate reports in any of the supported formats. Whereas, in **GroupDocs.Assembly for .NET 1.3.0** to generate a report in any of the supported format a separate namespace of that format is required to be added, given are the namespaces:

1.  using Groupdocs.Assembly;
2.  using Groupdocs.Assembly.Cells;
3.  using Groupdocs.Assembly.Data;
4.  using Groupdocs.Assembly.Exceptions;
5.  using Groupdocs.Assembly.Pdf;
6.  using Groupdocs.Assembly.Slides;
7.  using Groupdocs.Assembly.Words;

### Support of Multiple Data Sources

**GroupDocs.Assembly for .NET 3.0.0** supports following data sources:

*   Database
*   XML
*   JSON
*   OData
*   Custom .NET Objects

### Easy Template Syntax

**Next Generation GroupDocs.Assembly for .NET** engine supports the underlying C# syntax for LINQ queries which is actually shorter than LINQ. This basically means that the developers can use the familiar and well documented C# syntax to write data binding/traversal queries right in the document templates. As a result, developers can enjoy many benefits including short and concise reporting syntax and binding to any type of supported data source including business objects.

### Template Editing

In **GroupDocs.Assembly for .NET 1.3.0** user need to edit the template through code. However, template editing in **GroupDocs.Assembly for .NET 3.0.0** is quite straightforward. Font size, text alignment and other similar context is dealt on the template end.

### Comfy Code

Code to generate reports has been shrunk to fewer lines in **GroupDocs.Assembly for .NET 3.0.0**. A user does not need to write the bulk of lines of code to generate big reports. Instead, he/she will have to write 2 to 3 lines of code to generate business documents and reports.

{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/){{< /alert >}}

{{< gist GroupDocsGists 73103f412435effbc30b8f2c5b2366ba >}}