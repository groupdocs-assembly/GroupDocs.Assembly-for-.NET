---
id: numbered-list-in-email-document
url: assembly/net/numbered-list-in-email-document
title: Numbered List in Email Document
weight: 6
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will use GroupDocs.Assembly to generate a Numbered List report in Email Document format.{{< /alert >}}{{< alert style="info" >}}The code uses some of the objects defined in [The Business Layer](https://docs.groupdocs.com/assembly/net/the-business-layer/).{{< /alert >}}{{< alert style="info" >}}This feature is supported by version 17.8.0 or greater.{{< /alert >}}

## Numbered List in Email Document

### Creating a Numbered List

Practising the following steps you can create Numbered List Template in MS Outlook 2013.

1.  Create a new Email.
2.  Write a sentence like "We provide support for the following clients:".
3.  Start numbered list.
4.  Save the template.

### Reporting Requirement

As a report developer, you are required to represent the following key requirements:

*   The report must be in .eml or .msg format.
*   It must add email recipient, CSS and subject of the email.
*   The report must show the products in the numbered list.

### Adding Syntax to be evaluated by GroupDocs.Assembly Engine

```csharp
We provide support for the following products:
<<foreach [in products]>><<[NumberOf()]>>. <<[ProductName]>>
<</foreach>>
```

{{< alert style="success" >}}For detailed technical information about syntax, expressions and report generation by the engine, please visit: [Working with GroupDocs.Assembly Engine](https://docs.groupdocs.com/assembly/net/working-with-groupdocs-assembly-engine/).{{< /alert >}}

### Download Numbered List Template

Please download the sample Numbered List document we created in this article:

*   [Numbered List.msg](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/raw/master/Examples/Data/Source/Email%20Templates/Numbered%20List.msg?raw=true)

### Generating The Report

{{< gist GroupDocsGists a41628c2748f0237a84b710d9538fc4f >}}

## Numbering Restart in Nested Numbered List 

{{< alert style="info" >}}This feature is supported by version 19.7 or greater.{{< /alert >}}


The GroupDocs.Assembly engine allows restart list numbering within your documents dynamically using *<<restartNum>>* tags. In particular, this feature is useful when working with a nested numbered list within a data band.

Assume that we are picking *Order* and *Service* classes as defined in the following *Custom Objects* of our business use case.

{{< gist GroupDocsGists cd8768711ff88414192c3f50a9b7c918 >}}

Given that orders is an enumeration of Order instances, you could try to use the following template to output information on several orders in one document.

```csharp
<<foreach [in orders]>><<[Customer.CustomerName]>> (<<[Customer.CustomerContactNumber]>>)
1.	<<foreach [in Services]>><<[ServiceName]>>
<</foreach>><</foreach>>
```

The generated report will look as follows:

```
Jane Doe (+9211874)
	1.	Regular Cleaning
	2.	Oven Cleaning
John Smith (+458789)
	3.	Regular Cleaning
	4.	Oven Cleaning
	5.	Carpet Cleaning
John Smith (+458789)
	6.	Regular Cleaning
	7.	Carpet Cleaning
John Smith (+458789)
	8.	Oven Cleaning
```

However, there would be a single numbered list across all orders, which is not applicable for this scenario. Hence, you can make the list numbering to restart for every order by putting a *<<restartNum>>* tag into your template before a corresponding *<<foreach>>* tag as follows:

```csharp
<<foreach [in orders]>><<[ Customer.CustomerName]>> (<<[Customer.CustomerContactNumber]>>)
1.	<<restartNum>><<foreach [in Services]>><<[ServiceName]>>
<</foreach>><</foreach>>
```

 Then, the generated report will look as follows:

```
Jane Doe (+9211874)
	1.	Regular Cleaning
	2.	Oven Cleaning
John Smith (+458789)
	1.	Regular Cleaning
	2.	Oven Cleaning
	3.	Carpet Cleaning
John Smith (+458789)
	1.	Regular Cleaning
	2.	Carpet Cleaning
John Smith (+458789)
	1.	Oven Cleaning
```

## Download Numbered List Template

Please download the sample Numbered List template we created in this article:

*   [Numbered List\_RestartNum.msg](attachments/50266283/85426182.msg)
