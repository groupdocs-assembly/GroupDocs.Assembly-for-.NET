---
id: groupdocs-assembly-for-net-19-3-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-3-release-notes
title: GroupDocs.Assembly for .NET 19.3 Release Notes
weight: 7
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.3.{{< /alert >}}

## Major Features

Supported in-lining of syntax error messages into templates instead of exception throwing.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-99  | Provide an ability to inline template syntax errors in generated Word Processing documents instead of exception throwing  | Feature  |
| ASSEMBLYNET-100  | Provide an ability to inline template syntax errors in generated Spreadsheet documents instead of exception throwing  | Feature  |
| ASSEMBLYNET-101  | Provide an ability to inline template syntax errors in generated Presentation documents instead of exception throwing  | Feature  |
| ASSEMBLYNET-102  | Provide an ability to inline template syntax errors in generated Email documents instead of exception throwing  | Feature  |
| ASSEMBLYNET-103  | Provide an ability to inline template syntax errors in generated Plain Text documents instead of exception throwing  | Feature  |
| ASSEMBLYNET-107  | Allow null values for image tag expressions  | Enhancement  |
| ASSEMBLYNET-108  | Allow null values for doc tag expressions  | Enhancement  |

  

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.3. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported in-lining of syntax error messages into templates instead of exception throwing

By default, Document Assembler throws an exception when encounters a template syntax error. Such an exception provides information on a reason of the error and specifies a tag or expression part where the error is encountered. In most cases, this information is enough to find a place in a template causing the error and fix it.

However, when dealing with complex templates containing a large number of tags, it becomes harder to find an exact place in a template causing an error. To make things easier, the engine supports the *DocumentAssemblyOptions.InlineErrorMessages* option that enables in-lining of a syntax error message into a template document at an exact position where the error occurs during run-time.

{{< alert style="warning" >}}A template syntax error message is written using a bold font to make it more apparent.{{< /alert >}}

Consider the following template.

```csharp
<<var [name]>>
```

By default, such a template causes the engine to throw an exception while building a report. However, when **DocumentAssemblyOptions.InlineErrorMessages** is applied, no exception is thrown and the report looks as follows then.

```csharp
<<var [name] Error! An assignment operator is expected. >>

```

{{< alert style="warning" >}}Only messages describing errors in template syntax can be in-lined; messages describing errors encountered during expressions’ evaluation cannot.{{< /alert >}}

When* DocumentAssemblyOptions.InlineErrorMessages * is applied, a Boolean value returned by a *DocumentAssembler.AssembleDocument*overload indicates whether building of a report was finished successfully or was interrupted because of a template syntax error. This enables you to process reports which building succeeded or failed differently as shown in the following code snippet.

```csharp
DocumentAssembler assembler = new DocumentAssembler();

assembler.Options |= DocumentAssemblyOptions.InlineErrorMessages;

if (assembler.AssembleDocument (...))
{
    // Do something with a successfully
built report.
}
else
{
    // Do
something with a report containing a template syntax error.
}
```

{{< alert style="warning" >}}When DocumentAssemblyOptions.InlineErrorMessages is not applied, DocumentAssembler.AssembleDocument overloads return true if there were no template syntax errors encountered or throw an exception otherwise.{{< /alert >}}
