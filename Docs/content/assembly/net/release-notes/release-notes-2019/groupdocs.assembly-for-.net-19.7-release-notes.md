---
id: groupdocs-assembly-for-net-19-7-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-7-release-notes
title: GroupDocs.Assembly for .NET 19.7 Release Notes
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.7{{< /alert >}}

## Major Features

Supported dynamic list numbering restart for Word Processing documents and emails with HTML and RTF bodies.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-116  | Support dynamic list numbering restart for Word Processing documents  | Feature  |
| ASSEMBLYNET-118  | Support dynamic list numbering restart for emails with HTML and RTF bodies  | Feature  |
| ASSEMBLYNET-119  | Support credit-based metered licensing  | Feature  |

## Public API and Backward Incompatible Changes 

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.7. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Supported credit-based metered licensing

From now on, when using metered licensing, GroupDocs.Assembly consumes one credit (along with consumed bytes as previously) for a single document assembly operation.

### Added a new method to Metered

The following method was added.

```csharp
/// <summary>
/// Returns the currently consumed number of credits.
/// </summary>
/// <returns>The currently consumed number of credits.</returns>
public static decimal GetConsumptionCredit();
```

### Supported dynamic list numbering restart 

You can restart list numbering within your documents dynamically using restartNum tags. In particular, this feature is useful when working with a nested numbered list within a data band as shown in the following example.

Assume that you have the Order and Service classes defined in your application as follows.

```csharp
public class Order
{
    public String ClientName { get { ... } }
    public String ClientAddress { get { ... } }
    public IEnumerable<Service> Services { get { ... } }
    ...
}
 
public class Service
{
    public String Name { get { ... } }
    ...
}
```

Given that orders is an enumeration of Order instances, you could try to use the following template to output information on several orders in one document.

```csharp
<<foreach [order in orders]>><<[order.ClientName]>> (<<[order.ClientAddress]>>)
1. <<foreach [service in order.Services]>><<[service.Name]>>
<</foreach>><</foreach>>
```

But then, a result document would look as follows.

Jane Doe (445 Mount Eden Road Mount Eden Auckland 1024)
     1. Regular Cleaning
     2. Oven Cleaning
John Smith (43 Vogel Street Roslyn Palmerston North 4414)
     3. Regular Cleaning
     4. Oven Cleaning
     5. Carpet Cleaning

That is, there would be a single numbered list across all orders, which is not applicable for this scenario. However, you can make list numbering to restart for every order by putting a restartNum tag into your template before a corresponding foreach tag as follows.

```csharp
<<foreach [order in orders]>><<[order.ClientName]>> (<<[order.ClientAddress]>>)
1. <<restartNum>><<foreach [service in order.Services]>><<[service.Name]>>
<</foreach>><</foreach>>
```

{{< alert style="warning" >}}When using with a data band, it is required to put a restartNum tag before a corresponding foreach tag in the same numbered paragraph.{{< /alert >}}

Then, a result document looks as follows.

Jane Doe (445 Mount Eden Road Mount Eden Auckland 1024)

1.  Regular Cleaning
2.  Oven Cleaning

John Smith (43 Vogel Street Roslyn Palmerston North 4414)

1.  Regular Cleaning
2.  Oven Cleaning
3.  Carpet Cleaning
    

{{< alert style="warning" >}}You can use a restartNum tag without a data band to dynamically restart list numbering for a containing paragraph, if needed; for example, the tag can be used to restart list numbering for a document inserted dynamically.{{< /alert >}}
