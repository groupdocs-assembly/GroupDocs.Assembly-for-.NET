---
id: removing-empty-paragraphs
url: assembly/net/removing-empty-paragraphs
title: Removing Empty Paragraphs
weight: 39
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 18.10. or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}

GroupDocs.Assembly for .NET API supports the removal of paragraphs becoming empty after template syntax tags are removed or replaced with empty values. A new member **RemoveEmptyParagraphs** is added to [DocumentAssemblyOptions](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassemblyoptions)**.** When this new option is applied to [DocumentAssembler](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassembler) options through [DocumentAssembler.Options](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassembler/properties/options), the engine additionally removes empty paragraphs.

## Removing Empty Paragraphs in Word Processing Document

{{< gist GroupDocsGists 1baca62f7f2fb6b15feacc8686d431fa >}}

### Download

*   [Empty Paragraph.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Empty%20Paragraph.docx)

## Removing Empty Paragraphs in Presentation Document

{{< gist GroupDocsGists 7c9fc9e774e865e047e668ee915fea7e >}}

### Download

*   [Empty Paragraph.pptx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Presentation%20Templates/Empty%20Paragraph.pptx)

## Removing Empty Paragraphs in Email Document

{{< gist GroupDocsGists 5ba8a40ae2ca44cc4bd4576d5a2648e9 >}}

### Download

*   [Empty Paragraph.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Email%20Templates/Empty%20Paragraph.msg)