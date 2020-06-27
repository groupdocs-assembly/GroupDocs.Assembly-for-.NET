---
id: working-with-groupdocs-assembly-engine
url: assembly/net/working-with-groupdocs-assembly-engine
title: Working with GroupDocs.Assembly Engine
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026664064 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026664064 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026664064 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026664064"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#WorkingwithGroupDocs.AssemblyEngine-Aboutthissection">About this section</a></li><li><span class="TOCOutline">2</span> <a href="#WorkingwithGroupDocs.AssemblyEngine-Ataglance">At a glance</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1</span> <a href="#WorkingwithGroupDocs.AssemblyEngine-Stepsyoufollow">Steps you follow</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1.1</span> <a href="#WorkingwithGroupDocs.AssemblyEngine-TheactionsbyGroupDocs.AssemblyEngine">The actions by GroupDocs.Assembly Engine</a></li></ul></li><li><span class="TOCOutline">2.2</span> <a href="#WorkingwithGroupDocs.AssemblyEngine-Articlesinthissection">Articles in this section</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

## About this section

The articles in this section will be talking about technical concepts to be considered when creating document templates in order to generate your required reports. After going through this section, hopefully you will understand the template syntax, composing expressions and other C# syntax supported programming elements. Additionally, we'll also discuss how [GroupDocs.Assembly](https://apireference.groupdocs.com/net/assembly) engine reads the syntax, evaluates the expressions to populate with assigned data and then generates your desired report according to your business needs.

## At a glance

### Steps you follow

In a nutshell, you will follow these steps when using [GroupDocs.Assembly](https://apireference.groupdocs.com/net/assembly) Engine:

1.  Create a data source in one of the supported data formats.
2.  Create a template conforming to the supported syntax and expressions.
3.  Pass the data source and the template to the Engine.

#### The actions by GroupDocs.Assembly Engine

As a result of passing template and data source, the Engine will perform following actions:

1.  Sequentially evaluate the expressions against the passed data source object.
2.  Process the results of the expressions according to their roles.
3.  Replace the corresponding tags with appropriate contents.

### Articles in this section
