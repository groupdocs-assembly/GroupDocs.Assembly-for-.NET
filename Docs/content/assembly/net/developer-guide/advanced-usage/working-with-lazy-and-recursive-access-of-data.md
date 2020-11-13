---
id: working-with-lazy-and-recursive-access-of-data
url: assembly/net/working-with-lazy-and-recursive-access-of-data
title: Working with Lazy And Recursive Access of Data
weight: 26
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

## Purpose

At some point, you may want to or trying to use IDataRecord and IDataReader recursively in your application using GroupDocs.Assembly for .NET. But these Interfaces cannot serve the purpose. However, it is quite possible to accomplish the same goal using custom types.
See IDataReader and IDataRecords Implementors [here](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources).

## Creating Template

![](https://raw.githubusercontent.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/master/Examples/Data/Screenshots/template.PNG)

## Download Template

Get template from here.

*   [Lazy and recursive.docs](https://github.com/atirtahirgroupdocs/GroupDocs_Assembly_NET/blob/master/Examples/Data/Source/Word%20Templates/Lazy%20And%20Recursive.docx?raw=true)

## Generating The Report

{{< gist GroupDocsGists 96a8b32f705ac6c4ecd9cd2ebb698372 >}}


