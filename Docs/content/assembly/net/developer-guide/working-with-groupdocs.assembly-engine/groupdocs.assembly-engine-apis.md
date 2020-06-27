---
id: groupdocs-assembly-engine-apis
url: assembly/net/groupdocs-assembly-engine-apis
title: GroupDocs.Assembly Engine APIs
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}In this article, we will discuss the behavior of GroupDocs.Assembly APIs.{{< /alert >}}

<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026664469 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026664469 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026664469 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026664469"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#GroupDocs.AssemblyEngineAPIs-OverviewoftheAPI">Overview of the API</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#GroupDocs.AssemblyEngineAPIs-BuildingReports">Building Reports</a></li><li><span class="TOCOutline">1.2</span> <a href="#GroupDocs.AssemblyEngineAPIs-SettingupKnownExternalTypes">Setting up Known External Types</a></li><li><span class="TOCOutline">1.3</span> <a href="#GroupDocs.AssemblyEngineAPIs-OptimizingReflectionCalls">Optimizing Reflection Calls</a></li></ul></li><li><span class="TOCOutline">2</span> <a href="#GroupDocs.AssemblyEngineAPIs-TechnicalConsiderations">Technical Considerations</a><ul class="toc-indentation"><li><span class="TOCOutline">2.1</span> <a href="#GroupDocs.AssemblyEngineAPIs-ImplicitEnumerationDetermination">Implicit Enumeration Determination</a></li><li><span class="TOCOutline">2.2</span> <a href="#GroupDocs.AssemblyEngineAPIs-SupportforLINQProviders">Support for LINQ Providers</a></li><li><span class="TOCOutline">2.3</span> <a href="#GroupDocs.AssemblyEngineAPIs-MergingGroupDocs.Assembly">Merging GroupDocs.Assembly</a></li></ul></li></ul></div></div></div></td><td valign="top" width="15%">&nbsp;</td><td valign="top" width="35%">&nbsp;</td></tr></tbody></table>

## Overview of the API

Classes of GroupDocs.Assembly are located within the GroupDocs.Assembly namespace. The main of the classes is DocumentAssembler. All the functionality required to build a report from a template is assembled within the class.

### Building Reports

To build a report from a template, you can use one of the `ReportingEngine.BuildReport` overloads. The following table describes parameters of these overloads.

| Parameter | Description |
| --- | --- |
| document | A template document. At runtime, this document instance is populated with a data from the specified source and becomes a ready report. |
| dataSource | An object providing a data to populate the specified template. The object must be of one of the following types:
*   A traditional data source: [DataSet](http://msdn.microsoft.com/en-us/library/system.data.dataset(v=vs.110).aspx), [DataTable](http://msdn.microsoft.com/en-us/library/system.data.datatable(v=vs.110).aspx)
*   XML
*   OData
*   JSON
*   An object of a custom visible type: [Using Types](Template%2BSyntax%2B-%2BPart%2B1%2Bof%2B2.html#TemplateSyntax-Part1of2-UsingTypes){{< alert style="success" >}}For more details on supported Data Formats and Data Manipulation capabilities, please visit Data Sources.{{< /alert >}}

 |
| dataSourceName | The identifier of the specified data source object within the specified template. You can omit this identifier, if the template uses the [contextual object member access](Template%2BSyntax%2B-%2BPart%2B1%2Bof%2B2.html#TemplateSyntax-Part1of2-UsingContextualObjectMemberAccess) when dealing with the data source. |

Given a template to be populated with a data from a DataSet instance that is identified as "ds" within the template, you can use the following code to build the corresponding report.

```csharp
//Setting up source document template
const String strDocumentTemplate = "Template Folder/wordTemplate.docx";
//Setting up destination document report
const String strDocumentReport = "Report Folder/wordReport.docx";
// Setting up a data set.
DataSet ds = ........

DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(strDocumentTemplate, strDocumentReport, new DataSourceInfo(ds, "ds"));

```

Given a visible `Person` class defined in your application and a template to be populated with a data about a single Person instance using the contextual object member access, you can use the following code to build the corresponding report.

```csharp
//Setting up source document template
const String strDocumentTemplate = "Template Folder/wordTemplate.docx";
//Setting up destination document report
const String strDocumentReport = "Report Folder/wordReport.docx";
// Setting up a person data
Person person = ........

DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(strDocumentTemplate, strDocumentReport, new DataSourceInfo(person, "PersonDS"));

```

### Setting up Known External Types

GroupDocs.Assembly Engine must be aware of custom external types that you reference in your template before the engine processes the template. You can set up external types known by the engine through the ReportingEngine.KnownTypes property. The property represents an unordered set (that is, a collection of unique items) of [Type](http://msdn.microsoft.com/en-us/library/system.type(v=vs.110).aspx) objects. Every type in the set must meet requirements declared at [Working with Types](Template%2BSyntax%2B-%2BPart%2B1%2Bof%2B2.html#TemplateSyntax-Part1of2-UsingTypes).

**Note:** Aliases of simple types like int, string, and others are known by the engine by default.

Consider the following example. Given an ImageUtil class declared at your application and a template accessing a static member of this class, you can use the following code to make the engine be aware of the class before processing the template.

```csharp
DocumentAssembler assembler = new DocumentAssembler;
assembler.KnownTypes.Add(typeof(ImageUtil));
assembler.AssembleDocument(...);

```

### Optimizing Reflection Calls

GroupDocs.Assembly Engine uses reflection calls while accessing members of custom external types. However, reflection calls are much slower than direct calls, which creates a performance overhead.

Due to the performance overhead, the engine provides a strategy minimizing the reflection usage. The strategy is based upon the runtime type generation. That is, the engine generates a proxy type per an external type. The proxy directly calls members of the corresponding external type, the engine to access these members in a uniform way with no reflection involved. The proxy is [lazily initialized](http://en.wikipedia.org/wiki/Lazy_initialization) and reused in further. Thus, the reflection is used only while building the proxy.

Although this strategy can significantly minimize the reflection usage in a long run, it creates a performance overhead of the runtime type generation. So, if you deal with small data collections all the time while building your reports, consider the disabling of the strategy.

You can control the enabling of the strategy through the `DocumentAssembler.UseReflectionOptimization` static property. By default, the strategy is enabled.

## Technical Considerations

Here, we will reveal some technical aspects and implementation details related to the GroupDocs.Assembly Engine which can be useful for you while making design decisions for your applications.

### Implicit Enumeration Determination

If you do not specify the type of an enumeration item in a `foreach` statement or lambda function signature within your template explicitly, the type is implicitly determined by the engine from the type of the enumeration as follows:

1.  If the enumeration represents a `DataTable` instance, then the item type is `DataRow`.
2.  Otherwise, if the enumeration represents a `DataView` instance, then the item type is `DataRowView`.
3.  Otherwise, if the enumeration implements `IDataReader`, then the item type is `IDataRecord`.
4.  Otherwise, if the enumeration represents child rows of a `DataRow` instance, then the item type is `DataRow`.
5.  Otherwise, if the enumeration implements generic `IEnumerable<T>`, then the item type is a type argument corresponding to T. Note, that the engine does not support enumerations that implement `IEnumerable<T>` multiple times using different type arguments.
6.  Otherwise, the item type is `Object`.

### Support for LINQ Providers

The engine does not cooperate with [LINQ providers](http://msdn.microsoft.com/en-us/library/vstudio/bb882640(v=vs.110).aspx). To be aware of consequences of this feature, consider the following example.

Assume, that you have the `City` and `Person` classes defined in your application as follows.

```csharp
public class City
{
    public IList<Person> Persons { get { ... } }
    ...
}
public class Person
{
    public String Name { get { ... } }
    public int Age { get { ... } }
    ...
    }
}

```

Given that `city` is an instance of the `City` class, consider the evaluation of the following template expression by the engine.

```csharp
city.Persons.Where(p => p.Age > 30).OrderBy(p => p.Name)
```

If you use an ORM to provide an interaction of your application with an SQL database in conjunction with a LINQ provider for the ORM, then you most likely expect an SQL query like the following one to be performed in the background while evaluating the template expression.

```csharp
SELECT *
    FROM Persons
    WHERE CityId = @CityId AND Age > 30
    ORDER BY Name

```

However, since the engine does not cooperate with LINQ providers, the actual SQL query in this case is as follows.

```csharp
SELECT *
    FROM Persons
    WHERE CityId = @CityId

```

That is, the engine makes an eager call before applying its built-in extension methods. In some cases, this feature can lead to a significant performance overhead. So, if the performance of your application is your primary concern, then consider a preparing of your sequential data outside your templates. That is, pass to the engine a sequential data that is already filtered, ordered, grouped, and so forth.

### Merging GroupDocs.Assembly

You can use a tool like [ILMerge](http://research.microsoft.com/en-us/people/mbarnett/ilmerge.aspx) to merge GroupDocs.Assembly library with another assembly. But in order to use the Engine, please add the following attribute to the target assembly.

```csharp
[assembly: InternalsVisibleTo(
    "TargetAssemblyName" +
    ".ReportingDynamics" +
    ", PublicKey=TargetAssemblyPublicKey")]

```

`TargetAssemblyName` and `TargetAssemblyPublicKey` stand for the name of your target assembly and its public key (not a public key token) respectively.

Despite of the `InternalsVisibleTo` attribute applied, you can normally use obfuscation routines on your target assembly, if needed. An obfuscation does not harm a functionality of GroupDocs.Assembly Engine.
