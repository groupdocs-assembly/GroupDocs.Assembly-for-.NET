---
id: groupdocs-assembly-for-net-17-1-0-release-notes
url: assembly/net/groupdocs-assembly-for-net-17-1-0-release-notes
title: GroupDocs.Assembly for .NET 17.1.0 Release Notes
weight: 7
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This page contains release notes for GroupDocs.Assembly for .NET 17.1.0{{< /alert >}}

## Major Features

Extended abilities of working with tables of external documents as data sources for all supported file formats.

## Full List of Issues Covering all Changes in this Release

| Key | Summary | Category |
| --- | --- | --- |
| ASSEMBLYNET-26 | Add ability to load multiple DocumentTable objects from a single file as a single operation | Feature |
| ASSEMBLYNET-27 | Add ability to define relations between DocumentTable instances loaded from a single document | Feature |
| ASSEMBLYNET-28 | Add ability to change automatically detected types of DocumentTable columns | Feature |
| ASSEMBLYNET-29 | Add ability to use template syntax for related DataTable objects regardless of relation direction | Enhancement |

## Public API and Backward Incompatible Changes

{{< alert style="info" >}}This section lists public API changes that were introduced in GroupDocs.Assembly for .NET 17.1.0 It includes not only new and obsoleted public methods, but also a description of any changes in the behavior behind the scenes in GroupDocs.Assembly which may affect existing code. Any behavior introduced that could be seen as a regression and modifies existing behavior is especially important and is documented here.{{< /alert >}}

### Ability to load multiple DocumentTable objects from a single file as a single operation

The following classes and interfaces of the **GroupDocs.Assembly.Data** namespace have been added: **DocumentTableSet**, **DocumentTableCollection**, **IDocumentTableLoadHandler**, and **DocumentTableLoadArgs**.

The following properties of the **GroupDocs.Assembly.Data.DocumentTable** class has been added: **Name** and **IndexInDocument**.

**Example 1: Loading DocumentTableSet Using Default Options**

**C#**

```csharp
// Load all document tables using default options.
DocumentTableSet tableSet = new DocumentTableSet("Multiple Tables Data.docx");
 
// Check loading.
Debug.Assert(tableSet.Tables.Count == 3);
Debug.Assert(tableSet.Tables[0].Name == "Table1");
Debug.Assert(tableSet.Tables[1].Name == "Table2");
Debug.Assert(tableSet.Tables[2].Name == "Table3");
```

******Example 2: Loading DocumentTableSet Using Custom Options******

**C#**

```csharp
// Load document tables using custom options.
DocumentTableSet tableSet = new DocumentTableSet("Multiple Tables Data.docx", new CustomDocumentTableLoadHandler());
 
// Ensure that the second table is not loaded.
Debug.Assert(tableSet.Tables.Count == 2);
Debug.Assert(tableSet.Tables[0].Name == "Table1");
Debug.Assert(tableSet.Tables[1].Name == "Table3");
 
// Ensure that default options are used to load the first table (that is, default column names are used).
Debug.Assert(tableSet.Tables[0].Columns.Count == 2);
Debug.Assert(tableSet.Tables[0].Columns[0].Name == "Column1");
Debug.Assert(tableSet.Tables[0].Columns[1].Name == "Column2");
 
// Ensure that custom options are used to load the third table (that is, column names are extracted).
Debug.Assert(tableSet.Tables[1].Columns.Count == 2);
Debug.Assert(tableSet.Tables[1].Columns[0].Name == "Name");
Debug.Assert(tableSet.Tables[1].Columns[1].Name == "Address");
 
// ...
 
private class CustomDocumentTableLoadHandler : IDocumentTableLoadHandler
{
    public void Handle(DocumentTableLoadArgs args)
    {
        switch (args.TableIndex)
        {
            case 0:
                // Do nothing. The table is to be loaded with default options.
                break;
            case 1:
                // Discard loading of the table completely.
                args.IsLoaded = false;
                break;
            case 2:
                // Load the table with custom options.
                args.Options = new DocumentTableOptions();
                args.Options.FirstRowContainsColumnNames = true;
                break;
            default:
                throw new InvalidOperationException();
        }
    }
}
```

**Example 3: Using DocumentTableSet as Data Source**

**C#**

```csharp
// Set table column names to be extracted from the document.
DocumentTableSet tableSet = new DocumentTableSet("Multiple Tables Data.docx", new ColumnNameExtractingDocumentTableLoadHandler());
 
// Set table names for conveniency.
tableSet.Tables[0].Name = "Planets";
tableSet.Tables[1].Name = "Persons";
tableSet.Tables[2].Name = "Companies";
 
// Pass DocumentTableSet as a data source.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument("Using Document Table Set as Data Source.pptx", "Out.pptx", tableSet);
 
// ...
 
private class ColumnNameExtractingDocumentTableLoadHandler : IDocumentTableLoadHandler
{
    public void Handle(DocumentTableLoadArgs args)
    {
        args.Options = new DocumentTableOptions();
        args.Options.FirstRowContainsColumnNames = true;
    }
}
```

### Ability to define relations between DocumentTable instances loaded from a single document

The following classes of the **GroupDocs.Assembly.Data** namespace have been added: **DocumentTableRelation** and **DocumentTableRelationCollection**.

The **Relations** property of the **GroupDocs.Assembly.Data.DocumentTableSet** class has been added.

**Example: Using Document Table Relations**

**C#**

```csharp
// Set table column names to be extracted from the document.
DocumentTableSet tableSet = new DocumentTableSet("Related Tables Data.xlsx", new ColumnNameExtractingDocumentTableLoadHandler());
 
// Define relations between tables.
// NOTE: For Spreadsheet documents, table names are extracted from sheet names.
tableSet.Relations.Add(
    tableSet.Tables["CLIENT"].Columns["ID"],
    tableSet.Tables["CONTRACT"].Columns["CLIENT_ID"]);
 
tableSet.Relations.Add(
    tableSet.Tables["MANAGER"].Columns["ID"],
    tableSet.Tables["CONTRACT"].Columns["MANAGER_ID"]);
 
// Pass DocumentTableSet as a data source.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument("Using Document Table Relations.docx", "Out.docx", tableSet);
 
// ...
 
private class ColumnNameExtractingDocumentTableLoadHandler : IDocumentTableLoadHandler
{
    public void Handle(DocumentTableLoadArgs args)
    {
        args.Options = new DocumentTableOptions();
        args.Options.FirstRowContainsColumnNames = true;
    }
}
```

### Ability to change automatically detected types of DocumentTable columns

The **Type** property setter has been added to the **GroupDocs.Assembly.Data.DocumentTableColumn** class.

**Example: Changing Document Table Column Type**

**C#**

```csharp
// Set table column names to be extracted from the document.
DocumentTableOptions options = new DocumentTableOptions();
options.FirstRowContainsColumnNames = true;
 
DocumentTable table = new DocumentTable("Managers Data.docx", 1, options);
 
// NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
Debug.Assert(table.Columns["Total_Contract_Price"].Type == typeof(string));
 
// Change the column's type to double thus enabling to use arithmetic operations on values of the column 
// such as summing in templates.
table.Columns["Total_Contract_Price"].Type = typeof(double);
 
// Pass DocumentTable as a data source.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument("Changing Document Table Column Type.pptx", "Out.pptx", table, "Managers");
```
