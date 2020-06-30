---
id: template-syntax-part-1-of-2
url: assembly/net/template-syntax-part-1-of-2
title: Template Syntax - Part 1 of 2
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This article is the first part of the Template Syntax series of articles. For second part, please visit Template Syntax - Part 2 of 2.{{< /alert >}}

## Composing Template

A typical template for GroupDocs.Assembly Engine is composed of common document contents and tags that describe the template's structure and data bindings. You can form these tags using just running text that can occupy multiple paragraphs to be more descriptive.

A tag body must meet the following requirements:

*   A tag body must be surrounded by "<<" and ">>" character sequences.
*   A tag body must contain only text nodes.
*   A tag body must not be located inside markup document nodes.

A tag body typically consists of the following elements:

*   A tag name
*   An expression surrounded by brackets
*   A set of switches available for the tag, each of which is preceded by the "-" character

```csharp
<<tag_name [expression] -switch1 -switch2 ...>>
```

An optional comment can be written to provide a human-readable explanation.

```csharp
<<tag_name [expression] –switch1 –switch2 ... // optional_comment >>
```

{{< alert style="warning" >}}The optional comment feature is supported by version 19.1 or greater{{< /alert >}}

Particular tags can have additional elements. Some tags require closing counterparts. A closing tag has the "/" character that precedes its name. This tag's name must match to the name of the corresponding opening tag.

```csharp
<</tag_name>>
```

**Note:** Tag body elements are case-sensitive.

## Composing Expressions

### Using Lexical Tokens

The following table describes lexical tokens that you can use in template expressions and restrictions on these tokens' usage comparing with [C# Language Specification 5.0](http://www.microsoft.com/en-us/download/details.aspx?id=7029).

| Token | Restrictions |
| --- | --- |
| **Keyword** | Only the following tokens are reserved as keywords: `true`, `false`, `null`, `new`, and `in` |
| **Identifier** | 
*   The feature of keyword escaping through the "@" character is not supported.
*   Unicode character escapes are not permitted in identifiers.

 |
| **Literal** | None |
| **Operator** | See [Using Operators]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-1-of-2.md" >}}) |

You can use the following identifiers that are not preceded by a member access operator in template expressions:

*   The name of a passed data source object
*   The name of an iteration variable within its scope (see [Outputting Sequential Data]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#outputting-sequential-data" >}}) for more information)
*   The name of a lambda function parameter within the scope of the lambda function
*   A fully or partially qualified name of a type that is known by the engine (see [Setting up Known External Types]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/groupdocs.assembly-engine-apis.md#settingup-known-external-types" >}}) for more information)
*   The name of a member of an object that is determined as follows:
    *   Inside a data band body, the object is resolved to the innermost iteration variable.
    *   Outside a data band body, the object is resolved to a passed data source.

The feature of the omitting of an object identifier while accessing the object's members is also known as the contextual object member access. See [Using Contextual Object Member Access]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-1-of-2.md" >}}) for more information.

### Using Types

GroupDocs.Assembly Engine enables you to use external visible types in template expressions. A `visible type` is a public type with outer types (if any) are public as well. You can use a data source object of any visible type to pass it to the engine.

However, you can use the identifier of a visible type in template expressions only if the following additional requirements are met:

*   The type is not void.
*   The type does not represent an array.
*   The type is not an open or closed generic type.

**Note:** Whereas using of generic types' identifiers is forbidden in template expressions, you can use identifiers of nullable types in the shorthand `"T?"` form.

Also, the engine enables you to use anonymous types in template expressions. Such types are useful while composing expressions with grouping by multiple keys. See [Enumeration Extension Methods]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#enumeration-extension-methods" >}}) for the examples.

#### Type Members

GroupDocs.Assembly Engine enables you to access the following public (static and instance) members of accessible types in template expressions:

*   Fields
*   Properties
*   Indexer getters
*   Methods
*   Constructors

However, you can use a functional type member in template expressions only if the following additional requirements are met:

*   The functional member returns a value.
*   The functional member does not have `ref` or `out` parameters.
*   The functional member does not take generic type arguments.

The engine supports the following features when dealing with functional members:

*   Overload resolution according to [C# Language Specification 5.0](http://www.microsoft.com/en-us/download/details.aspx?id=7029)
*   Using of default parameter values without support of named parameters
*   Using of parameters taking a variable number of arguments

### Using Extension Methods

GroupDocs.Assembly Engine enables you to use the following built-in extension methods in template expressions:

*   Extension methods mimicking the ones for `IEnumerable<T>` (see [Enumeration Extension Methods]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#enumeration-extension-methods" >}}) for more information)
*   Extension methods for iteration variables (see [Extension Methods of Iteration Variables]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#extension-methodsof-iteration-variables" >}}) for more information)

**Note:** Extension methods, other than the built-in ones, can be used only in the form of plain static methods in template expressions.

### Using Operators

The following list contains predefined and user-defined operators that Document Assembler enables you to use in template expressions.

*   **Primary:**
    
    ```csharp
    x.y f(x) a[x] new x?.y a?[x]  
    ```
    
*   **Unary:**
    
    ```csharp
    - ! ~ (T)x
    ```
    
*   **Binary:**
    
    ```csharp
    * / % + - << >> < > <= >= == != & ^ | && || ??
    ```
    
*   **Ternary:**
    
    ```csharp
    ?:
    ```
    

The engine follows operator precedence, associativity and overload resolution rules declared at [C# Language Specification 5.0](http://www.microsoft.com/en-us/download/details.aspx?id=7029) while evaluating template expressions. But be aware of the following limitations in the behavior comparing with the specification:

*   Implicit user-defined conversions are supported only when specified explicitly.
*   The indexing of multi-dimensional arrays is not supported.
*   Whereas the object initializer syntax is supported (including objects of anonymous types), the collection initializer syntax is not.

Also, the engine enables you to use lifted operators in template expressions.

### Using Lambda Functions

GroupDocs.Assembly Engine enables you to use lambda functions only as arguments of built-in enumeration extension methods in template expressions. See [Enumeration Extension Methods]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#enumeration-extension-methods" >}}) for more information.

**Note:** Lambda functions declared within template expressions are not interchangeable with delegates. Thus, you can not pass delegates as arguments to built-in enumeration extension methods.

You can use both explicit and implicit lambda function signatures in template expressions. If you do not specify the type of a parameter of a lambda function explicitly, the type is determined implicitly by the engine depending on the type of the corresponding enumeration.

### Using Data Sources

#### DataSet Objects

{{< alert style="info" >}}This link will explain you how we utilized the DataSet Objects in Business Layer.{{< /alert >}}

GroupDocs.Assembly Engine enables you to access `DataTable` objects contained within a particular `DataSet` instance by table names using the "." operator in template expressions. That is, for example, given that `ds` is a `DataSet` instance that contains a `DataTable` named "Persons", you can access the table using the following syntax.

```csharp
ds.Persons

```

**Note** Table names are case-insensitive.

#### DataTable and DataView Objects

GroupDocs.Assembly Engine enables you to treat `DataTable` and `DataView` objects in template expressions as enumerations of their rows. That is, you can use template expressions evaluated to such objects in `foreach` tags (see [Outputting Sequential Data]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#outputting-sequential-data" >}}) for more information).

Also, you can normally apply enumeration extension methods (see [Enumeration Extension Methods]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#enumeration-extension-methods" >}}) for more information) to `DataTable` and `DataView` objects in template expressions. For example, given that persons is a `DataTable` or `DataView` instance, you can count its rows using the following syntax.

```csharp
persons.Count()

```

#### DataRow and DataRowView Objects

GroupDocs.Assembly Engine enables you to access a data associated with a particular DataRow or DataRowView instance in template expressions using the "." operator. The following table describes, which identifiers you can use to access different kinds of the data.

| Data Kind | Identifier | Examples of Template Expressions |
| --- | --- | --- |
| **Field Value** | Field name | Given that r is a row that has a field named "Name", you can access the field’s value using the following syntax.
```csharp
r.Name
```





 |
| **Single Parent or Child Row** | Parent (child) table name | Given that r is a row of a *DataTable* that has a parent (child) *DataTable* named "City", you can access the single parent (child) row of r using the following syntax.

```csharp
r.City
```

Given that the "City" *DataTable* has a field named "Name", you can access the field’s value for the single parent (child) row using the following syntax.

```csharp
r.City.Name
```

 |
| **Enumeration of Child or Parent Rows** | Child (parent) table name | Given that r is a row of a *DataTable* that has a child (parent) *DataTable* named "Persons", you can access the enumeration of the child (parent) rows of r using the following syntax.

```csharp
r.Persons
```

Given that the "Persons" *DataTable* has a field named "Age", you can count the child (parent) rows that correspond to persons over thirty years old using the following syntax.

```csharp
r.Persons.Count(p => p.Age > 30)
```

 |

{{< alert style="warning" >}} Field and table names are case-insensitive.{{< /alert >}}

To determine parent-child relationships for a particular *DataTable* instance, the engine uses *[DataRelation](http://msdn.microsoft.com/en-us/library/system.data.datarelation(v=vs.110).aspx)* objects contained within the corresponding *DataSet* instance. Thus, you can manage these relationships in a common way.

{{< alert style="warning" >}}Instead of using of table names to access data of child or parent rows, you can also use relation names, which is useful when you deal with multiple relations to the same table.{{< /alert >}}

  

#### IDataReader Implementors

GroupDocs.Assembly Engine enables you to treat IDataReader implementors as enumerations of IDataRecord implementors in template expressions. That is, you can use IDataReader implementors in template expressions in the same way as DataTable objects. See [Working with DataTable and DataView Objects]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-1-of-2.md" >}}) for more information.

However, you can not use IDataReader implementors in template expressions in conjunction with enumeration operations that require a caching of enumeration items. The examples of such operations are grouping and sorting. To work around this restriction, use DataTable objects instead.

#### IDataRecord Implementors

{{< alert style="info" >}}See In-Table Master-Detail template using IDataRecord Implementors.{{< /alert >}}

Document Assembler enables you to access field values of a particular `IDataRecord` implementor by field names using the "." operator in template expressions. To use this feature, one of the following conditions must be met:

*   The `IDataRecord` implementor represents an iteration variable upon enumerating an `IDataReader` implementor (see [Outputting Sequential Data]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-2-of-2.md#outputting-sequential-data" >}}) for more information).
*   The `IDataRecord` implementor does not implement the `IDataReader` interface.

The following example shows, how to use this feature. Given that r is an `IDataRecord` implementor that has a field named "Name", you can access the field's value using the following syntax.

```csharp
r.Name
```

**Note** Field names are case-insensitive.

### Using Images

{{< alert style="info" >}}See Common List with Image template using images in it.{{< /alert >}}

You can insert images to your reports dynamically using image tags. To declare a dynamically inserted image within your template, please follow these steps:

1.  Add a textbox to your template at the place where you want an image to be inserted.
2.  Set common image attributes such as frame, size, and others for the textbox, making the textbox look like a blank inserted image.
3.  Specify an image tag within the textbox using the following syntax.
    
    ```csharp
    <<image [image_expression]>>
    
    ```
    

The expression declared within an image tag is used by the engine to build an image to be inserted. The expression must return a value of one of the following types:

*   *   A byte array containing an image data
    *   A [Stream](http://msdn.microsoft.com/en-us/library/system.io.stream(v=vs.110).aspx) instance able to read an image data
    *   An [Image](http://msdn.microsoft.com/en-us/library/system.drawing.image(v=vs.110).aspx) object
    *   A string containing an image URI, path, or Base64-encoded image data

While building a report, the following procedure is applied to an image tag:

*   The expression declared within the tag is evaluated and its result is used to form an image.
*   The corresponding textbox is filled with this image.
*   The tag is removed from the textbox.

**Note:** If the expression declared within an image tag returns a stream object, then it is closed by the engine as soon as the corresponding image is built.

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
    

### Inserting Documents

{{< alert style="info" >}}See Inserting Documents Dynamically for more details.{{< /alert >}}

You can insert contents of outer documents to your reports dynamically using doc tags. A doc tag denotes a placeholder within a template for a document to be inserted during runtime.

Syntax of a doc tag is defined as follows.

```csharp
<<doc [document_expression]>>
```

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

### Setting Checkbox Values Dynamically

{{< alert style="info" >}}See Set Checkbox Values Dynamically for more details.{{< /alert >}}

You can set checkbox values to either checked or unchecked in your reports dynamically by taking the following steps:

1.  Add a checkbox content control to your template at a place where you want it to appear in a result document.
2.  By editing content control properties, add a check tag to the title of the checkbox content control using the following syntax.

```csharp
<<check [conditional_expression]>>
```

Here, conditional\_expression defines a condition upon which the value of the checkbox content control is to be set to checked (or unchecked, if the condition is not met). The conditional expression must return a Boolean value.

While building a report, a conditional expression declared within a check tag is evaluated and the value of its containing checkbox content control is set accordingly. The check tag itself is removed then.

{{< alert style="warning" >}}Dynamic checkbox value setting is available for Word Processing documents only.{{< /alert >}}

### Using Hyperlinks 

Using GroupDocs.Assembly, you can insert hyperlinks to URI or Bookmarks to your reports dynamically using link tags. The syntax of a link tag is defined below for various types of documents:

#### Word Processing and Emails

```csharp
<<link [uri_or_bookmark_expression][display_text_expression]>>
```

#### Spreadsheets

If the insertion of the link to cell A1 is required:

```csharp
<<link ["A1"] ["Home"]>>
```

#### Presentations

```csharp
<<link ["Slide1"] ["Home"]>>
```

### Using Contextual Object Member Access

{{< alert style="info" >}}See In-Table List with Filtering, Grouping, and Ordering template Contextual Object Member Access in it.{{< /alert >}}

You can make your templates less cumbersome using the contextual object member access feature. This feature enables you to access members of some objects without specifying the objects' identifiers in template expressions. An object to which the feature can be applied is determined depending on a context as follows:

*   Inside a data band body, the object is resolved to the innermost iteration variable.
*   Outside a data band body, the object is resolved to a passed data source.

Obviously, inside a data band body, you can not use the feature to access members of an outer iteration variable or a passed data source object. With the exception of this restriction, you can use both contextual and common object member access syntaxes interchangeably depending on your needs and preferences.

Consider the following example. Given that `ds` is a `DataSet` instance containing a `DataTable` object named "Persons" that has fields named "Name" and "Age", you can use the following template to list the contents of the table.

| No. | Name | Age |
| --- | --- | --- |
| 
```csharp
<<foreach [p
        in ds.Persons]>><<[
    p.NumberOf()]>>
```



 | 

```csharp
<<[p.Name]>>
```



 | 

```csharp
<<[p.Age]>><</
foreach>>
```



 |
| 

```csharp
Count: <<[ds.Persons.Count()]>>
```



 | 

 | 

 |

Alternatively, you can use the following template involving the contextual object member access syntax to get the same results.

| No. | Name | Age |
| --- | --- | --- |
| 
```csharp
 <<foreach [
        in Persons]>><<[
    NumberOf()]>>
```



 | 

```csharp
<<[Name]>> 
```



 | 

```csharp
<<[Age]>><</
foreach>> 
```



 |
| 

```csharp
Count: <<[Persons.Count()]>> 
```



 | 

 | 

 |

### Using Conditional Blocks

You can use different document blocks to represent the same data depending on a condition with the help of conditional blocks. A conditional block represents a set of template options, each of which is bound with a conditional expression. At runtime, these conditional expressions are sequentially evaluated, until an expression that returns true is reached. Then, the conditional block is replaced with the corresponding template option populated with data.  
A conditional block can have a default template option that is not bound with a conditional expression. At runtime, this template option is used, when none of conditional expressions return true. If a default template option is missing and none of conditional expressions return true, then the whole conditional block is removed during runtime.

You can use the following syntax to declare a conditional block:

```csharp
<<if [conditional_expression1]>>
template_option1
<<elseif [conditional_expression2]>>
template_option2
...
<<else>>
default_template_option
<</if>>

```

**Note:** A conditional expression must return a Boolean value.

#### Common Conditional Blocks

{{< alert style="info" >}}See Multicolored Numbered List template using Common Conditional Blocks in it.{{< /alert >}}

A common conditional block is a conditional block which body starts and ends within paragraphs that belong to a single story or table cell.  
If a conditional block belongs to a single paragraph, it can be used as a replacement for an expression tag that involves the ternary "?:" operator. For example, given that items is an enumeration, you can use the following template to represent the count of elements in the enumeration:

```csharp
You have chosen <<if [!items.Any()]>>no items<<else>><<[items.Count()]>> item(s)<</if>>.

```

**Note:** A template option of a common conditional block can be composed of multiple paragraphs, if needed.

You can normally use common conditional blocks within data bands. For example, given that items is an enumeration of the strings "item1", "item2", and "item3", you can use the following template to enumerate them and apply different formatting for even and odd elements:

<<foreach \[item in items\]>><<if \[IndexOf() % 2 == 0\]>><<\[item\]>>  
<<else>><<\[item\]>>  
<</if>><</foreach>>

In this case, the engine produces a report as follows:

item1  
item2  
item3

You can use data bands within common conditional blocks as well. For example, given the previous declaration of items, you can check whether the enumeration contains any elements before outputting their list:

```csharp
<<if [!items.Any()]>>No data.
<<else>><<foreach [item in items]>><<[item]>>
<</foreach>><</if>>

```

#### Table-Row Conditional Block

{{< alert style="info" >}}See In-Table List with Highlighted Rows template using Table-Row Conditional Block.{{< /alert >}}

A table-row conditional block is a conditional block which body occupies single or multiple rows of a single document table. The body of such a block (as well as the body of its every template option) starts at the beginning of the first occupied row and ends at the end of the last occupied row as follows:  
**Note:** Table rows occupied by different template options in the following template are highlighted with different colors.

td { border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; }

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&nbsp;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&nbsp;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&nbsp;</span></p></td></tr><tr><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;if ...&gt;&gt; ...</span></p></td><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(255, 242, 204);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;elseif ...&gt;&gt; ...</span></p></td><td style="background-color: rgb(255, 242, 204);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(255, 242, 204);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(255, 242, 204);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(255, 242, 204);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(255, 242, 204);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt; ...</span></p></td><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(222, 234, 246);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">... &lt;&lt;/if&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&nbsp;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&nbsp;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&nbsp;</span></p></td></tr></tbody></table>

The following examples in this section are given using client, an instance of the Client class, and clients, an enumeration of instances of the Client class that is defined as follows:

```csharp
public class Client
{
    public String Name { get { ... } }
    public String Country { get { ... } }
    public String LocalAddress { get { ... } }
    ...
}

```

Using table-row conditional blocks, you can pick to output a single row among several rows of a single document table depending on a condition like in the following example:

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;if [client.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;"> == "New Zealand"]&gt;&gt;&lt;&lt;[client.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Name</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td colspan="2" style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[client.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">LocalAddress</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[client.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Name</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[client.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[client.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">LocalAddress</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;/if&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">...</span></p></td></tr><tr style="height: 0pt;"><td style="width: 176.95pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 127.6pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 162.7pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

You can normally use table-row conditional blocks within data bands to make elements of an enumeration look differently depending on a condition. Consider the following template:

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;foreach [in clients]&gt;&gt;&lt;&lt;if [</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;"> == "New Zealand"]&gt;&gt;&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Name</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td colspan="2" style="background-color: rgb(226, 239, 217);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">LocalAddress</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Name</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(242, 242, 242);"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">LocalAddress</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;&lt;&lt;/if&gt;&gt;&lt;&lt;/foreach&gt;&gt;</span></p></td></tr><tr style="height: 0pt;"><td style="width: 176.95pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 127.6pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 162.7pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

In this case, the engine produces the report as below:

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">A Company</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 116.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Australia</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 151.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">219-241 Cleveland St</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">STRAWBERRY HILLS&nbsp; NSW&nbsp; 1427</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">B Ltd.</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 116.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Brazil</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 151.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Avenida João Jorge, 112, ap. 31</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Vila Industrial</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Campinas - SP</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">13035-680</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">C &amp; D</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 116.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Canada</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 151.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">101-3485 RUE DE LA MONTAGNE</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">MONTRÉAL (QUÉBEC) h2G 2A6</span></p></td></tr><tr><td style="background-color: rgb(226, 239, 217); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">E Corp.</span></p></td><td colspan="2" style="background-color: rgb(226, 239, 217); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 279.5pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">445 Mount Eden Road</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Mount Eden</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Auckland 1024</span></p></td></tr><tr><td style="background-color: rgb(226, 239, 217); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">F &amp; Partners</span></p></td><td colspan="2" style="background-color: rgb(226, 239, 217); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 279.5pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">20 Greens Road</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Tuahiwi</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Kaiapoi 7691</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">G &amp; Co.</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 116.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Greece</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 151.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Karkisias 6</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">GR-111 42&nbsp; ATHINA</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">GRÉCE</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">H Group</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 116.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Hungary</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 151.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Budapest</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Fiktív utca 82., IV. em./28.</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">2806</span></p></td></tr><tr><td style="background-color: rgb(226, 239, 217); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">I &amp; Sons</span></p></td><td colspan="2" style="background-color: rgb(226, 239, 217); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 279.5pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">43 Vogel Street</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Roslyn</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Palmerston North 4414</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 166.15pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">J Ent.</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 116.8pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Japan</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 151.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Hakusan 4-Ch?me 3-2</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Bunky?-ku, T?KY?</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">112-0001</span></p><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Japan</span></p></td></tr><tr style="height: 0pt;"><td style="width: 176.95pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 127.6pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 162.7pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

**Note:** You can use common conditional blocks within table-row data bands as well.

Also, you can use data bands inside table-row conditional blocks. For example, you can provide an alternate content for an empty table-row data band using the following template:

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 144.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Client</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 102.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 187.35pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Local Address</span></p></td></tr><tr><td colspan="3" style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 456.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: center;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;if [!clients.</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Any</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">()]&gt;&gt;No data</span></p></td></tr><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 144.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;else&gt;&gt;&lt;&lt;foreach [in clients]&gt;&gt;&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Name</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 102.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 187.35pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">&lt;&lt;[</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">LocalAddress</span><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">]&gt;&gt;&lt;&lt;/foreach&gt;&gt;&lt;&lt;/if&gt;&gt;</span></p></td></tr><tr style="height: 0pt;"><td style="width: 155.7pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 113.4pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 198.15pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

In case, the corresponding enumeration is empty, the engine produces a report as below:

<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-left: 0pt;"><tbody><tr><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 144.9pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Client</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 102.6pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Country</span></p></td><td style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 187.35pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">Local Address</span></p></td></tr><tr><td colspan="3" style="background-color: rgb(242, 242, 242); border-bottom-color: rgb(0, 0, 0); border-bottom-style: solid; border-bottom-width: 0.75pt; border-left-color: rgb(0, 0, 0); border-left-style: solid; border-left-width: 0.75pt; border-right-color: rgb(0, 0, 0); border-right-style: solid; border-right-width: 0.75pt; border-top-color: rgb(0, 0, 0); border-top-style: solid; border-top-width: 0.75pt; padding-left: 5.03pt; padding-right: 5.03pt; vertical-align: top; width: 456.45pt;"><p style="margin-top: 0pt; margin-right: 0pt; margin-bottom: 0pt; margin-left: 0pt; text-align: center;"><span style="color: rgb(127, 127, 127); font-family: &quot;Courier New&quot;; font-size: 11pt; font-weight: bold;">No data</span></p></td></tr><tr style="height: 0pt;"><td style="width: 155.7pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 113.4pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td><td style="width: 198.15pt; border-top-color: initial; border-top-style: none; border-top-width: initial; border-right-color: initial; border-right-style: none; border-right-width: initial; border-bottom-color: initial; border-bottom-style: none; border-bottom-width: initial; border-left-color: initial; border-left-style: none; border-left-width: initial;"></td></tr></tbody></table>

**Note:** If tags denoting boundaries of a template option are contained within a single table cell, the option is considered to be a common template option, rather than a table-row one. That is, the option is considered to occupy contents within the cell, rather than the whole row. That is why, a single-cell alternate content in the previous example is located between an opening if and else tags, rather than between an else and closing if tags.
