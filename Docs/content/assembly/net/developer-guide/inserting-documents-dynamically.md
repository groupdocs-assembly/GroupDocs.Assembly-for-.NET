---
id: inserting-documents-dynamically
url: assembly/net/inserting-documents-dynamically
title: Inserting Documents Dynamically
weight: 43
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is supported by version 20.3 or greater.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in The Business Layer.{{< /alert >}}

You can insert contents of outer documents to your reports dynamically using doc tags. A doc tag denotes a placeholder within a template for a document to be inserted during run time.

Syntax of a doc tag is defined as follows.

```csharp
<<doc [document_expression]>>
```

And simply call the assembler method to generate report like following code snippets:

<table class="highlight tab-size js-file-line-container" data-tab-size="8" data-paste-markdown-skip=""><tbody><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L1" class="blob-num js-line-number" data-line-number="1"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span> For complete examples and data files, please go to https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET</span></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L2" class="blob-num js-line-number" data-line-number="2"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-k">try</span></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L3" class="blob-num js-line-number" data-line-number="3"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC3" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L4" class="blob-num js-line-number" data-line-number="4"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC4" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up source document template (Email or Word Document)</span></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L5" class="blob-num js-line-number" data-line-number="5"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentTemplate</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Templates/DynamicDocInsert.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L6" class="blob-num js-line-number" data-line-number="6"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC6" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L7" class="blob-num js-line-number" data-line-number="7"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC7" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Setting up destination for reports</span></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L8" class="blob-num js-line-number" data-line-number="8"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC8" class="blob-code blob-code-inner js-file-line"><span class="pl-k">const</span> <span class="pl-en">String</span> <span class="pl-smi">strDocumentReport</span> <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">"</span>Word Reports/DynamicDocInsert Out.docx<span class="pl-pds">"</span></span>;</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L9" class="blob-num js-line-number" data-line-number="9"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC9" class="blob-code blob-code-inner js-file-line"></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L10" class="blob-num js-line-number" data-line-number="10"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC10" class="blob-code blob-code-inner js-file-line"><span class="pl-c"><span class="pl-c">//</span>Assemble the document</span></td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L11" class="blob-num js-line-number" data-line-number="11"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC11" class="blob-code blob-code-inner js-file-line"><span class="pl-en">DocumentAssembler</span> <span class="pl-smi">assembler</span> <span class="pl-k">=</span> <span class="pl-k">new</span> <span class="pl-en">DocumentAssembler</span>();</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L12" class="blob-num js-line-number" data-line-number="12"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC12" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">assembler</span>.<span class="pl-en">AssembleDocument</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetSourceDocument</span>(<span class="pl-smi">strDocumentTemplate</span>),</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L13" class="blob-num js-line-number" data-line-number="13"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC13" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">CommonUtilities</span>.<span class="pl-en">SetDestinationDocument</span>(<span class="pl-smi">strDocumentReport</span>), <span class="pl-k">new</span> <span class="pl-en">DataSourceInfo</span>(<span class="pl-smi">CommonUtilities</span>.<span class="pl-en">GetOuterDocumentFolder</span>(<span class="pl-s"><span class="pl-pds">"</span>OuterDocument.docx<span class="pl-pds">"</span></span>), <span class="pl-s"><span class="pl-pds">"</span>document_expression<span class="pl-pds">"</span></span>));</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L14" class="blob-num js-line-number" data-line-number="14"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC14" class="blob-code blob-code-inner js-file-line">}</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L15" class="blob-num js-line-number" data-line-number="15"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">catch</span> (<span class="pl-en">Exception</span> <span class="pl-smi">ex</span>)</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L16" class="blob-num js-line-number" data-line-number="16"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC16" class="blob-code blob-code-inner js-file-line">{</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L17" class="blob-num js-line-number" data-line-number="17"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-smi">Console</span>.<span class="pl-en">WriteLine</span>(<span class="pl-smi">ex</span>.<span class="pl-smi">Message</span>);</td></tr><tr><td id="file-insertdocumentdynamicallyinword_20-3-cs-L18" class="blob-num js-line-number" data-line-number="18"></td><td id="file-insertdocumentdynamicallyinword_20-3-cs-LC18" class="blob-code blob-code-inner js-file-line">}</td></tr></tbody></table>

[view raw](https://gist.github.com/GroupDocsGists/b06d6c5655c84fc772c6411d66016943/raw/0e3f7c8194319630dfb51af12d10c100c846cbcb/InsertDocumentDynamicallyInWord_20.3.cs) [InsertDocumentDynamicallyInWord\_20.3.cs](https://gist.github.com/GroupDocsGists/b06d6c5655c84fc772c6411d66016943#file-insertdocumentdynamicallyinword_20-3-cs) hosted with ❤ by [GitHub](https://github.com)

{{< alert style="warning" >}}A doc tag can be used almost anywhere in a template document except textboxes and charts.{{< /alert >}}

An expression declared within a doc tag is used by the assembler to load a document to be inserted during runtime. The expression must return a value of one of the following types:

*   A byte array containing document data
*   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read document data
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
