---
id: working-with-string-template
url: assembly/net/working-with-string-template
title: Working with String Template
weight: 37
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026667118 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026667118 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026667118 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026667118"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#WorkingwithStringTemplate-WorkingwithStringTemplate">Working with String Template</a></li><li><span class="TOCOutline">2</span> <a href="#WorkingwithStringTemplate-TheCode">The Code</a></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

## Working with String Template

GroupDocs.Assembly API allows you to use a string as input and output template, instead of document. This feature improves programmability while working with strings instead of documents, the resultant output string can be used as per the application needs. Follow these steps to using a string as input and output template:

*   Initialize [DocumentAssembler](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassembler)
*   Define **sourceString @"<<\[yourValue\]>>"**
*   Convert sourceString to **sourceBytes** using **Encoding.UTF8.GetBytes(sourceString)**
*   Get**targetBytes** using [DocumentAssembler.AssembleDocument()](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly/documentassembler/methods/assembledocument/index)
*   Get **targetString** using **Encoding.UTF8.GetString(targetBytes)**

## The Code
