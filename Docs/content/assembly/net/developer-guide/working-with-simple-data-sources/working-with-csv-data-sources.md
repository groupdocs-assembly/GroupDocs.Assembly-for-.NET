---
id: working-with-csv-data-sources
url: assembly/net/working-with-csv-data-sources
title: Working with CSV Data Sources
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 19.10 or later releases.{{< /alert >}}

To access CSV data while building a report, you can pass a [CsvDataSource](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/csvdatasource)instance to the assembler as a data source. Using of *CsvDataSource* enables you to work with typed values rather than just strings in template documents. Although CSV as a format does not define a way to store values of types other than strings, *CsvDataSource* is capable to recognize values of the following types by their string representations:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?

{{< alert style="warning" >}}For recognition of data types to work, string representations of corresponding values must be formed using invariant culture settings.{{< /alert >}}

### Treating simple CSV data

In template documents, a [CsvDataSource ](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/csvdatasource)instance should be treated in the same way as if it was a *DataTable* instance (see "[Working with *DataTable* and *DataView* Objects]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-1-of-2.md#datatable-and-dataview-objects" >}})" for more information) as shown in the following example.

Suppose we have CSV data like:
```
John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm
```

By using the template like following:
```
<<foreach \[in persons\]>>Name: <<\[Column1\]>>, Age: <<\[Column2\]>>, Date
of Birth: <<\[Column3\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Column2)\]>>
```

The results will be produced like:
```
Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36
```

{{< alert style="warning" >}}Using of the custom date-time format and the extension method involving arithmetic in the template document becomes possible, because text values of Column3 and Column2 are automatically converted to DateTime? and Int32? respectively.{{< /alert >}}

### Configure data source to read column names

By default, [CsvDataSource ](https://apireference.groupdocs.com/net/assembly/groupdocs.assembly.data/csvdatasource)uses column names such as "*Column1*", "*Column2*", and so on, as you can see from the previous example. However, *CsvDataSource* can be configured to read column names from the first line of CSV data as shown in the following example.

Suppose we have CSV data like:
```
Name,Age,Birth
John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm
```

By using the template like following:
```
<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of
Birth: <<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>
```

The code should be written like:
{{< gist GroupDocsGists f59b71ca7e6345f9ac817297f3707f5f >}}

The results will be produced like:
```
Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36
```

Also, you can use *CsvDataLoadOptions* to customize the following characters playing special roles while loading CSV data:

*   Value separator (the default is comma)
*   Single-line comment start (the default is sharp)
*   Quotation mark enabling to use other special characters within a value (the default is double quotes)

###  Download

#### Data Source Document

*   [Person.csv](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/XML%20DataSource/Managers.xml?raw=true)

#### Template

*   [Demo Template.txt](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)
