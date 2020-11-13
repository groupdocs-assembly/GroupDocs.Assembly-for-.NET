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
