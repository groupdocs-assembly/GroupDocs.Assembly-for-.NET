---
id: working-with-xml-data-sources
url: assembly/net/working-with-xml-data-sources
title: Working with XML Data Sources
weight: 3
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 19.10 or later releases.{{< /alert >}}

To access XML data while building a report, you can use facilities of *DataSet* to read XML into it and then pass it to the assembler as a data source. However, if your scenario does not permit to specify XML schema while loading XML into *DataSet*, all attributes and text values of XML elements are loaded as strings then. Thus, it becomes impossible, for example, to use arithmetic operations on numbers or to specify custom date-time and numeric formats to output orresponding values, because all of them are treated as strings. To overcome this limitation, you can pass an *XmlDataSource* instance to the assembler as a data source instead. Even when XML schema is not provided, *XmlDataSource* is capable to recognize values of the following types by their string representations:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?

{{< alert style="warning" >}}For recognition of data types to work, string representations of corresponding attributes and text values of XML elements must be formed using invariant culture settings.{{< /alert >}}

### Treating the top-level XML element contains only a sequence of elements of the same type

In template documents, if a top-level XML element contains only a sequence of elements of the same type, an *XmlDataSource* instance should be treated in the same way as if it was a *DataTable *instance (see "[Working with *DataTable* and *DataView* Objects]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-1-of-2.md#datatable-and-dataview-objects" >}})" for more information).

Suppose we have XML data like:

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

By using the template like following:

  

<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth: 
<<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>

  

The results will be produced like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36

{{< alert style="warning" >}}Using of the custom date-time format and the extension method involving arithmetic in the template document becomes possible, because text values of Birth and Age XML elements are automatically converted to DateTime? and Int32? respectively even in the absence of XML schema.{{< /alert >}}

### Treating the top-level XML element contains attributes or nested elements of different types

If a top-level XML element contains attributes or nested elements of different types, an *XmlDataSource* instance should be treated in template documents in the same way as if it was a *DataRow* instance (see "[Working with *DataRow* and *DataRowView* Objects]({{< ref "assembly/net/developer-guide/working-with-groupdocs.assembly-engine/template-syntax-part-1-of-2.md#datatable-and-dataview-objects" >}})" for more information) as shown in the following example.

Suppose we have XML data like:

<Person>
   <Name>John Doe</Name>
   <Age>30</Age>
   <Birth>1989-04-01 4:00:00 pm</Birth>
   <Child>Ann Doe</Child>
   <Child>Charles Doe</Child>
</Person>

By using the template like following:

Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth:
<<\[Birth\]:"dd.MM.yyyy">>
Children:
<<foreach \[in Child\]>><<\[Child\_Text\]>>
<</foreach>>

The results will be produced like:

Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Children:
Ann Doe
Charles Doe

{{< alert style="warning" >}}To reference a sequence of repeated simple-type XML elements with the same name, the elements’ name itself (for example, "Child") should be used in a template document, whereas the same name with the "_Text" suffix (for example, "Child_Text") should be used to reference the text value of one of these elements.{{< /alert >}}

### The complete example

The following example sums up typical scenarios involving nested complex-type XML elements.

#### XML

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

{{< gist GroupDocsGists ddac17d9a69c71495b2931ca6c7e8ce7 >}}



  

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

### Download

#### Data Source Document

*   [Managers.xml](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/XML%20DataSource/Managers.xml?raw=true)

#### Template

*   [Demo Template.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)
