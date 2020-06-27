---
id: groupdocs-assembly-for-net-17-5-0-release-notes
url: assembly/net/groupdocs-assembly-for-net-17-5-0-release-notes
title: GroupDocs.Assembly for .NET 17.5.0 Release Notes
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 17.5.0{{< /alert >}}

## Major Features

Support of variables, dynamic text background setting, and a new image size fit mode are added.

## Full List of Issues Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-37 | Add ability to define and use variables in template documents | Feature |
| ASSEMBLYNET-38 | Add ability to set text background color dynamically | Feature |
| ASSEMBLYNET-39 | Add ability not to grow the size of an image container while fitting it for the size of an image | Feature |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 17.5.0 It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}{{< alert style="warning" >}}Dynamic setting of a background color through the backColor tag is not supported for HTML, XML, and Plain Text file formats. However, there is an alternative approach to achieve this for HTML documents as described in the following article.{{< /alert >}}

### Setting Background Color Dynamically for HTML Documents

For HTML documents, a background color can not be set dynamically using the backColor tag. However, you can use the HTML style attribute to achieve this as shown in the following examples.

### Example 1. Setting a background color for a text fragment dynamically

Given that color is the string "orange", you can use the following template to highlight a text fragment dynamically by applying the corresponding background color to it.

```csharp
<html>
<body>
Regular text. <span style="background-color:&lt;&lt;[color]&gt;&gt;">Highlighted text.</span>
</body>
</html>

```

In this case, GroupDocs.Assembly produces the following result document.

```csharp
<html>
<body>
Regular text. <span style="background-color:orange">Highlighted text.</span>
</body>
</html>
```

### Example 2. Setting background colors for table rows dynamically

Given that colors is an enumeration of the strings "red", "orange", and "yellow", you can use the following template to apply corresponding background colors to rows of a table.

```csharp
<html>
<body>
<table>
&lt;&lt;foreach [color in colors]&gt;&gt;
<tr style="background-color:&lt;&lt;[color]&gt;&gt;">
<td>
Item
</td>
<td>
Some text
</td>
</tr>
&lt;&lt;/foreach&gt;&gt;
</table>
</body>
</html>
```

In this case, GroupDocs.Assembly produces the following result document.

```csharp
<html>
<body>
<table>

<tr style="background-color:red">
<td>
Item
</td>
<td>
Some text
</td>
</tr>

<tr style="background-color:orange">
<td>
Item
</td>
<td>
Some text
</td>
</tr>

<tr style="background-color:yellow">
<td>
Item
</td>
<td>
Some text
</td>
</tr>

</table>
</body>
</html>

```
