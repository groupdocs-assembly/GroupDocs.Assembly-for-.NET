---
id: working-with-json-data-sources
url: assembly/net/working-with-json-data-sources
title: Working with JSON Data Sources
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 19.10 or later releases.{{< /alert >}}

To access JSON data while building a report, the GroupDocs.Assembly API introduces *JsonDataSource* class. You can pass its instance to the assembler as a data source. Using of *JsonDataSource* enables you to work with typed values of JSON elements in template

documents. For more convenience, the set of simple JSON types is extended as follows:

*   Int32?
*   Int64?
*   Double?
*   Boolean?
*   DateTime?
*   String

### Treating the top level arrays or objects having array

In template documents, if a top-level JSON element is an array or an object having only one property of an array type, a JsonDataSource instance should be treated in the same way as if it was a DataTable instance (see "[Working with DataTable and DataView Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information) .

Suppose we have Json data like:
```
[ 
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
]
```

or alternative JSON like:
```
{ 
   Persons:[ 
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
   ]
}
```

If we use the template like following:
```
<<foreach \[in persons\]>>Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth: <<\[Birth\]:"dd.MM.yyyy">>
<</foreach>>
Average age: <<\[persons.Average(p => p.Age)\]>>
```

The results will be produced for both pieces of JSON data like:
```
Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Name: Jane Doe, Age: 27, Date of Birth: 31.01.1992
Name: John Smith, Age: 51, Date of Birth: 08.03.1968
Average age: 36
```

{{< alert style="warning" >}}Using of the custom date-time format becomes possible, because text values of Birth properties are automatically converted to DateTime?.{{< /alert >}}

### Treating the objects at top level

If a top-level JSON element represents an object, a *JsonDataSource* instance should be treated in template documents in the same way as if it was a *DataRow* instance (see "[Working with *DataRow* and *DataRowView* Objects](https://docs.groupdocs.com/assembly/net/template-syntax-part-1-of-2/#using-data-sources)" for more information). If a top-level JSON object has a single property that is also an object, then this nested object is accessed by the assembler instead.

Suppose we have Json data like:
```
{ 
   Name:"John Doe",
   Age:30,
   Birth:"1989-04-01 4:00:00 pm",
   Child:[ 
      "Ann Doe",
      "Charles Doe"
   ]
}
```

or alternatively like:
```
{ 
   Person:{ 
      Name:"John Doe",
      Age:30,
      Birth:"1989-04-01 4:00:00 pm",
      Child:[ 
         "Ann Doe",
         "Charles Doe"
      ]
   }
}
```

If we use the template like following:
```
Name: <<\[Name\]>>, Age: <<\[Age\]>>, Date of Birth:
<<\[Birth\]:"dd.MM.yyyy">>
Children:
<<foreach \[in Child\]>><<\[Child\_Text\]>>
<</foreach>>
```

The results will be produced for both pieces of JSON data like:
```
Name: John Doe, Age: 30, Date of Birth: 01.04.1989
Children:
Ann Doe
Charles Doe
```

{{< alert style="warning" >}}To reference a JSON object property that is an array of simple-type values, the name of the property (for example, "Child") should be used in a template document, whereas the same name with the "_Text" suffix (for example, "Child_Text") should be used to reference the value of an item of this array.{{< /alert >}}

### The complete example

The following example sums up typical scenarios involving nested JSON objects and arrays.

#### JSON
```
[ 
   { 
      Name:"John Smith",
      Contract:[ 
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
      ]
   },
   { 
      Name:"Tony Anderson",
      Contract:[ 
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
      ]
   },
   { 
      Name:"July James",
      Contract:[ 
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
      ]
   }
]
```

#### Template document
```
<<foreach [in managers]>>Manager: <<[Name]>>
Contracts:
<<foreach [in Contract]>>- <<[Client.Name]>> ($<<[Price]>>)
<</foreach>>
<</foreach>>
```

#### Source code

{{< gist GroupDocsGists 330892e76bb1958b9199e3d223962e27 >}}

#### Result document
```
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
```

### Download

#### Data Source Document

*   [ManagerData.json](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Data%20Sources/JSON%20DataSource/ManagerData.json?raw=true)

#### Template

*   [Demo Template.docx](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Source/Word%20Templates/Using%20Spreadsheet%20as%20Table%20of%20Data.docx?raw=true)
