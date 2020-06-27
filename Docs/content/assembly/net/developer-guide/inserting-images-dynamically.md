---
id: inserting-images-dynamically
url: assembly/net/inserting-images-dynamically
title: Inserting Images Dynamically
weight: 44
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 20.3 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

You can insert images to your reports dynamically using image tags. To declare a dynamically inserted image within your template, do the following steps:

1.  Add a textbox to your template at the place where you want an image to be inserted.
2.  Set common image attributes such as frame, size, and others for the textbox, making the textbox look like a blank inserted image.
3.  Specify an image tag within the textbox using the following syntax.

```csharp
<<image [image_expression]>>
```

And simply call the assembler method to generate report like following code snippets:

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC3" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source document template (Email or Word Document)</span></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Templates/DynamicImageDemo.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC6" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC7" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination for reports</span></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/DynamicImageDemo Out.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC9" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Assemble the document</span></td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>), <span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetImageFolder</span>(<span class="pl-s"><span class="pl-pds">"</span>no-photo.jpg<span class="pl-pds">"</span></span>),<span class="pl-s"><span class="pl-pds">"</span>image_expression<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC14" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC16" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-insertimagedynamicallyinword_20-3-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-insertimagedynamicallyinword_20-3-cs-LC18" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/b06d6c5655c84fc772c6411d66016943/raw/0e3f7c8194319630dfb51af12d10c100c846cbcb/InsertImageDynamicallyInWord_20.3.cs) [InsertImageDynamicallyInWord\_20.3.cs](https://gist.github.com/GroupDocsGists/b06d6c5655c84fc772c6411d66016943#file-insertimagedynamicallyinword_20-3-cs) hosted with ❤ by [GitHub](https://github.com)

  

The expression declared within an image tag is used by the assembler to build an image to be inserted. The expression must return a value of one of the following types:

*   A byte array containing an image data
*   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read an image data
*   An [Image](http://msdn.microsoft.com/en-us/library/system.drawing.image(v=vs.110).aspx) object
*   A string containing an image URI, path, or Base64-encoded image data

While building a report, the following procedure is applied to an image tag:

1.  The expression declared within the tag is evaluated and its result is used to form an image.
2.  The corresponding textbox is filled with this image.
3.  The tag is removed from the textbox.

{{< alert style="warning" >}}If the expression declared within an image tag returns a stream object, then it is closed by the assembler as soon as the corresponding image is built. {{< /alert >}}

By default, the assembler stretches an image filling a textbox to the size of the textbox. However, you can change this behavior in the following ways:

*   To keep the size of the textbox and stretch the image within bounds of the textbox preserving the ratio of the image, use the *keepRatio* switch as follows:
    
    ```csharp
    <<image [image_expression] -keepRatio>>
    ```
    
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
    
    {{< alert style="warning" >}}If the size of the image is greater than the size of the textbox, then the fitSizeLim switch acts like fitHeight or fitWidth. Otherwise, the fitSizeLim switch acts like fitSize.{{< /alert >}}
    

{{< alert style="warning" >}}Dynamic insertion of images from Base64-encoded bytes is available for file formats where dynamic image insertion is available for almost all supported file formats.{{< /alert >}}
