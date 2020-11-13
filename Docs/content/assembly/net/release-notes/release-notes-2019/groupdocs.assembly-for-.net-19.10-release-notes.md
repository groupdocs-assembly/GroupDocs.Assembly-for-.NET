---
id: groupdocs-assembly-for-net-19-10-release-notes
url: assembly/net/groupdocs-assembly-for-net-19-10-release-notes
title: GroupDocs.Assembly for .NET 19.10 Release Notes
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 19.10{{< /alert >}}

## Major Features

Simplified working with XML, JSON, and CSV data sources.

## Full List of Features Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-122  | Simplify working with XML data sources  | Feature  |
| ASSEMBLYNET-123  | Simplify working with JSON data sources  | Feature  |
| ASSEMBLYNET-124  | Support working with CSV data sources  | Feature  |
| ASSEMBLYNET-120  | List numbering restart changes list indent for emails with HTML bodies  | Bug  |

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 19.10. It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

## Public API and Backward Incompatible Changes 

### Simplified working with JSON data sources

To access JSON data while building a report, you can pass a *JsonDataSource* instance to the assembler as a data source. Using of *JsonDataSource* enables you to work with typed values of JSON elements in template

documents. For more convenience, the set of simple JSON types is extended as follows:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?
*   String

In template documents, if a top-level JSON element is an array or an object having only one property of an array type, a JsonDataSource instance should be treated in the same way as if it was a DataTable instance (see "[Working with DataTable and DataView Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information) as shown in the following example.

#### **JSON**

\[ 
   { 
      Name:"John Doe",
      Age:30,
      Birth:"1989-04-01 4:00:00 pm"
   },
   { 
      Name:"Jane Doe",
      Age:27,
      Birth:"1992-01-31 07:00:00 am"
   },
   { 
      Name:"John Smith",
      Age:51,
      Birth:"1968-03-08 1:00:00 pm"
   }
\]

#### Alternative JSON (produces the same result)

{ 
   Persons:\[ 
      { 
         Name:"John Doe",
         Age:30,
         Birth:"1989-04-01 4:00:00 pm"
      },
      { 
         Name:"Jane Doe",
         Age:27,
         Birth:"1992-01-31 07:00:00 am"
      },
      { 
         Name:"John Smith",
         Age:51,
         Birth:"1968-03-08 1:00:00 pm"
      }
   \]
}

#### Template document

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth: <<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
JsonDataSource dataSource = ... // Loading JSON.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource, "persons"));
```

#### Result document

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format becomes possible, because text values of Birth properties are automatically converted to DateTime?.{{< /alert >}}


If a top-level JSON element represents an object, a *JsonDataSource* instance should be treated in template documents in the same way as if it was a *DataRow* instance (see "[Working with *DataRow* and *DataRowView* Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information). If a top-level JSON object has a single property that is also an object, then this nested object is accessed by the assembler instead. To see how it works, consider the following example.

#### JSON

{ 
   Name:"John Doe",
   Age:30,
   Birth:"1989-04-01 4:00:00 pm",
   Child:\[ 
      "Ann Doe",
      "Charles Doe"
   \]
}

#### Alternative JSON (produces the same result)

{ 
   Person:{ 
      Name:"John Doe",
      Age:30,
      Birth:"1989-04-01 4:00:00 pm",
      Child:\[ 
         "Ann Doe",
         "Charles Doe"
      \]
   }
}

#### Template document

Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth:
<<\[Birth\]:"dd.MM.yyyy">>
Children:
<<foreach \[in Child\]>><<\[Child\_Text\]>>
<</foreach>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
JsonDataSource dataSource = ... // Loading JSON.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource));
```

#### Result document

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Children:
Ann Doe
Charles Doe

{{< alert style="warning" >}}To reference a JSON object property that is an array of simple-type values, the name of the property (for example, "Child") should be used in a template document, whereas the same name with the "_Text" suffix (for example, "Child_Text") should be used to reference the value of an item of this array.{{< /alert >}}

The following example sums up typical scenarios involving nested JSON objects and arrays.

#### JSON

\[ 
   { 
      Name:"John Smith",
      Contract:\[ 
         { 
            Client:{ 
               Name:"A Company"
            },
            Price:1200000
         },
         { 
            Client:{ 
               Name:"B Ltd."
            },
            Price:750000
         },
         { 
            Client:{ 
               Name:"C & D"
            },
            Price:350000
         }
      \]
   },
   { 
      Name:"Tony Anderson",
      Contract:\[ 
         { 
            Client:{ 
               Name:"E Corp."
            },
            Price:650000
         },
         { 
            Client:{ 
               Name:"F & Partners"
            },
            Price:550000
         }
      \]
   },
   { 
      Name:"July James",
      Contract:\[ 
         { 
            Client:{ 
               Name:"G & Co."
            },
            Price:350000
         },
         { 
            Client:{ 
               Name:"H Group"
            },
            Price:250000
         },
         { 
            Client:{ 
               Name:"I & Sons"
            },
            Price:100000
         },
         { 
            Client:{ 
               Name:"J Ent."
            },
            Price:100000
         }
      \]
   }
\]

#### Template document

<<foreach \[in managers\]>>Manager: <<\[Name\]>>
Contracts:
<<foreach \[in Contract\]>>- <<\[Client.Name\]>> ($<<\[Price\]>>)
<</foreach>>
<</foreach>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
JsonDataSource dataSource = ... // Loading JSON.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource, "managers"));
```

#### Result document

Manager: John Smith
Contracts:
- A Company ($1200000)
- B Ltd. ($750000)
- C & D ($350000)
Manager: Tony Anderson
Contracts:
- E Corp. ($650000)
- F & Partners ($550000)
Manager: July James
Contracts:
- G & Co. ($350000)
- H Group ($250000)
- I & Sons ($100000)
- J Ent. ($100000)

### Simplified working with XML data sources

To access XML data while building a report, you can use facilities of *DataSet* to read XML into it and then pass it to the assembler as a data source. However, if your scenario does not permit to specify XML schema while loading XML into *DataSet*, all attributes and text values of XML elements are loaded as strings then. Thus, it becomes impossible, for example, to use arithmetic operations on numbers or to specify custom date-time and numeric formats to output orresponding values, because all of them are treated as strings. To overcome this limitation, you can pass an *XmlDataSource* instance to the assembler as a data source instead. Even when XML schema is not provided, *XmlDataSource* is capable to recognize values of the following types by their string representations:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?

{{< alert style="warning" >}}For recognition of data types to work, string representations of corresponding attributes and text values of XML elements must be formed using invariant culture settings.{{< /alert >}}

In template documents, if a top-level XML element contains only a sequence of elements of the same type, an *XmlDataSource* instance should be treated in the same way as if it was a *DataTable *instance (see "[Working with *DataTable* and *DataView* Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information) as shown in the following example.

#### XML

<Persons>
   <Person>
      <Name>John Doe</Name>
      <Age>30</Age>
      <Birth>1989-04-01 4:00:00 pm</Birth>
   </Person>
   <Person>
      <Name>Jane Doe</Name>
      <Age>27</Age>
      <Birth>1992-01-31 07:00:00 am</Birth>
   </Person>
   <Person>
      <Name>John Smith</Name>
      <Age>51</Age>
      <Birth>1968-03-08 1:00:00 pm</Birth>
   </Person>
</Persons>

#### Template document

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date ofBirth: 
<<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
XmlDataSource dataSource = ... // Loading XML (without schema).
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource, "persons"));
```

#### Result document

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format and the extension method involving arithmetic in the template document becomes possible, because text values of Birth and Age XML elements are automatically converted to DateTime? and Int32? respectively even in the absence of XML schema.{{< /alert >}}

If a top-level XML element contains attributes or nested elements of different types, an *XmlDataSource* instance should be treated in template documents in the same way as if it was a *DataRow* instance (see "[Working with *DataRow* and *DataRowView* Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information) as shown in the following example.

#### XML

<Person>
   <Name>John Doe</Name>
   <Age>30</Age>
   <Birth>1989-04-01 4:00:00 pm</Birth>
   <Child>Ann Doe</Child>
   <Child>Charles Doe</Child>
</Person>

#### Template document

Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth:
<<\[Birth\]:"dd.MM.yyyy">>
Children:
<<foreach \[in Child\]>><<\[Child\_Text\]>>
<</foreach>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
XmlDataSource dataSource = ... // Loading XML.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource));
```

#### Result document

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Children:
Ann Doe
Charles Doe

{{< alert style="warning" >}}To reference a sequence of repeated simple-type XML elements with the same name, the elements’ name itself (for example, "Child") should be used in a template document, whereas the same name with the "_Text" suffix (for example, "Child_Text") should be used to reference the text value of one of these elements.{{< /alert >}}

The following example sums up typical scenarios involving nested complex-type XML elements.

### XML

<Managers>
   <Manager>
      <Name>John Smith</Name>
      <Contract>
         <Client>
            <Name>A Company</Name>
         </Client>
         <Price>1200000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>B Ltd.</Name>
         </Client>
         <Price>750000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>C &amp; D</Name>
         </Client>
         <Price>350000</Price>
      </Contract>
   </Manager>
   <Manager>
      <Name>Tony Anderson</Name>
      <Contract>
         <Client>
            <Name>E Corp.</Name>
         </Client>
         <Price>650000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>F &amp; Partners</Name>
         </Client>
         <Price>550000</Price>
      </Contract>
   </Manager>
   <Manager>
      <Name>July James</Name>
      <Contract>
         <Client>
            <Name>G &amp; Co.</Name>
         </Client>
         <Price>350000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>H Group</Name>
         </Client>
         <Price>250000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>I &amp; Sons</Name>
            43
         </Client>
         <Price>100000</Price>
      </Contract>
      <Contract>
         <Client>
            <Name>J Ent.</Name>
         </Client>
         <Price>100000</Price>
      </Contract>
   </Manager>
</Managers>

#### Template document

<<foreach \[in managers\]>>Manager: <<\[Name\]>>
Contracts:
<<foreach \[in Contract\]>>- <<\[Client.Name\]>> ($<<\[Price\]>>)
<</foreach>>
<</foreach>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
XmlDataSource dataSource = ... // Loading XML.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource, "managers"));
```

#### Result document

Manager: John Smith
Contracts:
- A Company ($1200000)
- B Ltd. ($750000)
- C & D ($350000)
Manager: Tony Anderson
Contracts:
- E Corp. ($650000)
- F & Partners ($550000)
Manager: July James
Contracts:
- G & Co. ($350000)
- H Group ($250000)
- I & Sons ($100000)
- J Ent. ($100000)

### Supported working with CSV data sources

To access CSV data while building a report, you can pass a *CsvDataSource* instance to the assembler as a data source. Using of *CsvDataSource* enables you to work with typed values rather than just strings in template documents. Although CSV as a format does not define a way to store values of types other than strings, *CsvDataSource* is capable to recognize values of the following types by their string representations:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?

{{< alert style="warning" >}}For recognition of data types to work, string representations of corresponding values must be formed using invariant culture settings.{{< /alert >}}

In template documents, a *CsvDataSource* instance should be treated in the same way as if it was a *DataTable* instance (see "[Working with *DataTable* and *DataView* Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information) as shown in the following example.

#### CSV

John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm

#### Template document

<<foreach \[in persons\]>>Name: <<\[Column1\]>>, Age: <<\[Column2\]>>, Date
of Birth: <<\[Column3\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Column2)\]>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
CsvDataSource dataSource = ... // Loading CSV.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource, "persons"));
```

Result document

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format and the extension method involving arithmetic in the template document becomes possible, because text values of Column3 and Column2 are automatically converted to DateTime? and Int32? respectively.{{< /alert >}}

By default, *CsvDataSource* uses column names such as "Column1", "Column2", and so on, as you can see from the previous example. However, *CsvDataSource* can be configured to read column names from the first line of CSV data as shown in the following example.

#### CSV

Name,Age,Birth
John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm

#### Template document

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of
Birth: <<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

#### Source code

```csharp
string template_doc_path = ... // Path for template doc
string report_doc_path = ... // path for report doc
CsvDataLoadOptions options = new CsvDataLoadOptions(true);
CsvDataSource dataSource = new CsvDataSource(..., options); // Loading CSV.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(template_doc_path, report_doc_path, new DataSourceInfo(dataSource, "persons"));
```

#### Result document

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

Also, you can use *CsvDataLoadOptions* to customize the following characters playing special roles while loading CSV data:

*   Value separator (the default is comma)
*   Single-line comment start (the default is sharp)
*   Quotation mark enabling to use other special characters within a value (the default is double quotes)
