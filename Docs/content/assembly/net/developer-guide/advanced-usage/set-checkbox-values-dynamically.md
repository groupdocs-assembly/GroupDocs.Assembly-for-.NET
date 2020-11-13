---
id: set-checkbox-values-dynamically
url: assembly/net/set-checkbox-values-dynamically
title: Set Checkbox Values Dynamically
weight: 45
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 20.3 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

You can set checkbox values to either checked or unchecked in your reports dynamically by taking the following steps:

1.  Add a checkbox content control to your template at a place where you want it to appear in a result document.
2.  By editing content control properties, add a check tag to the title of the checkbox content control using the following syntax.

```csharp
<<check [conditional_expression]>>
```

The code will look like following snippet:

{{< gist GroupDocsGists b06d6c5655c84fc772c6411d66016943 SetCheckboxValueDynamicallyInWord_20.3.cs >}}

Here, conditional\_expression defines a condition upon which the value of the checkbox content control is to be set to checked (or unchecked, if the condition is not met). The conditional expression must return a Boolean value.

While building a report, a conditional expression declared within a check tag is evaluated and the value of its containing checkbox content control is set accordingly. The check tag itself is removed then.

{{< alert style="warning" >}}Dynamic checkbox value setting is available for Word Processing documents only.{{< /alert >}}
