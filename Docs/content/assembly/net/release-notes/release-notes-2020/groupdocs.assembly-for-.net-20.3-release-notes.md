---
id: groupdocs-assembly-for-net-20-3-release-notes
url: assembly/net/groupdocs-assembly-for-net-20-3-release-notes
title: GroupDocs.Assembly for .NET 20.3 Release Notes
weight: 90
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
## Major Features

Supported dynamic insertion of documents and images from Base64-encoded bytes and dynamic checkbox value setting.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-145  | Support dynamic checkbox value setting for Word Processing documents  | Feature  |
| ASSEMBLYNET-147  | Support dynamic image and document insertion from Base64-encoded bytes  | Feature  |
| ASSEMBLYNET-148  | ArgumentException is thrown if image or document expression returns an empty string  | Bug  |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 20.3. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported dynamic insertion of documents and images from Base64-encoded bytes and dynamic checkbox value setting

#### Inserting Documents Dynamically

You can insert contents of outer documents to your reports dynamically using doc tags. A doc tag denotes a placeholder within a template for a document to be inserted during runtime.

Syntax of a doc tag is defined as follows.

```csharp
<<doc [document_expression]>>
```

{{< alert style="warning" >}}A doc tag can be used almost anywhere in a template document except textboxes and charts.{{< /alert >}}

An expression declared within a doc tag is used by the assembler to load a document to be inserted during runtime. The expression must return a value of one of the following types:

*   A byte array containing document data
*   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read document data
*   An instance of the Document class
*   A string containing a document URI, path, or Base64-encoded document data

While building a report, an expression declared within a doc tag is evaluated and its result is used to load a document which content replaces the doc tag then.

{{< alert style="warning" >}}If an expression declared within a doc tag returns a stream object, then the stream is closed by the assembler as soon as a corresponding document is loaded.{{< /alert >}}

By default, a document being inserted is not checked against template syntax and is not populated with data. However, you can enable this by using a build switch as follows.

```csharp
<<doc [document_expression] -build>>
```

When a build switch is used, the assembler treats a document being inserted as a template that can access the following data available at the scope of a corresponding doc tag:

*   Data sources
*   Variables
*   A contextual object 
*   Known external types 

{{< alert style="warning" >}}Dynamic insertion of documents from Base64-encoded bytes is available for file formats where dynamic document insertion is available for Word Processing documents and emails with HTML and RTF bodies only.{{< /alert >}}

#### Inserting Images Dynamically

You can insert images to your reports dynamically using image tags. To declare a dynamically inserted image within your template, do the following steps:

1.  Add a textbox to your template at the place where you want an image to be inserted.
2.  Set common image attributes such as frame, size, and others for the textbox, making the textbox look like a blank inserted image.
3.  Specify an image tag within the textbox using the following syntax.

```csharp
<<image [image_expression]>>
```

The expression declared within an image tag is used by the assembler to build an image to be inserted. The expression must return a value of one of the following types:

*   A byte array containing an image data
*   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read an image data
*   An [Image](http://msdn.microsoft.com/en-us/library/system.drawing.image(v=vs.110).aspx) object
*   A string containing an image URI, path, or Base64-encoded image data

While building a report, the following procedure is applied to an image tag:

1.  The expression declared within the tag is evaluated and its result is used to form an image.
2.  The corresponding textbox is filled with this image.
3.  The tag is removed from the textbox.

{{< alert style="warning" >}}If the expression declared within an image tag returns a stream object, then it is closed by the assembler as soon as the corresponding image is built. {{< /alert >}}

By default, the assembler stretches an image filling a textbox to the size of the textbox. However, you can change this behavior in the following ways:

*   To keep the width of the textbox and change its height preserving the ratio of the image, use the fitHeight switch as follows.

```csharp
<<image [image_expression] -fitHeight>>
```

*   To keep the height of the textbox and change its width preserving the ratio of the image, use the fitWidth switch as follows.

```csharp
<<image [image_expression] -fitWidth>>
```

*   To change the size of the textbox according to the size of the image, use the fitSize switch as follows.

```csharp
<<image [image_expression] -fitSize>>
```

*   To change the size of the textbox according to the size of the image without increasing the size of the textbox, use the fitSizeLim switch as follows.

```csharp
<<image [image_expression] -fitSizeLim>>
```

{{< alert style="warning" >}}If the size of the image is greater than the size of the textbox, then the fitSizeLim switch acts like fitHeight or fitWidth. Otherwise, the fitSizeLim switch acts like fitSize.{{< /alert >}}{{< alert style="warning" >}}Dynamic insertion of images from Base64-encoded bytes is available for file formats where dynamic image insertion is available for almost all supported file formats.{{< /alert >}}

#### Setting Checkbox Values Dynamically

You can set checkbox values to either checked or unchecked in your reports dynamically by taking the following steps:

1.  Add a checkbox content control to your template at a place where you want it to appear in a result document.
2.  By editing content control properties, add a check tag to the title of the checkbox content control using the following syntax.

```csharp
<<check [conditional_expression]>>
```

Here, conditional\_expression defines a condition upon which the value of the checkbox content control is to be set to checked (or unchecked, if the condition is not met). The conditional expression must return a Boolean value.

While building a report, a conditional expression declared within a check tag is evaluated and the value of its containing checkbox content control is set accordingly. The check tag itself is removed then.

{{< alert style="warning" >}}Dynamic checkbox value setting is available for Word Processing documents only.{{< /alert >}}
