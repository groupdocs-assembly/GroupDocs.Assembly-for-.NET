---
id: business-case-of-a-shopping-store
url: assembly/net/business-case-of-a-shopping-store
title: Business Case of a Shopping Store
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This business case has been conceived for the demonstration purpose and serves well for a typical shopping store. There may be simpler or complex scenarios in real life businesses.{{< /alert >}}

## Business Case

1.  There are many customers of the shopping store.
2.  The store offers multiple products.
3.  A customer may place multiple orders.
4.  An order may have many products.
5.  The data may or may not be stored in a database system.

## Reporting Requirements

1.  A report based on some criteria is to be generated.
2.  The report can be in any of the formats: Documents, Spreadsheets, Presentations.
3.  The report developer must be able to define templates and then generate reports.

## Business Objects

### Class Diagram

![](https://raw.githubusercontent.com/groupdocsassembly/GroupDocs_Assembly_NET/master/Examples/Data/Data%20Sources/Database%20Backup/class.PNG)

### Custom Objects

Here are our customer objects:

{{< gist GroupDocsGists cd8768711ff88414192c3f50a9b7c918 >}}



### Database Entities

![](https://raw.githubusercontent.com/groupdocsassembly/GroupDocs_Assembly_NET/master/Examples/Data/Data%20Sources/Database%20Backup/database.PNG)

## Templates Location

There is a fixed location of templates somewhere /Date/Source/. The report developer may create any number of templates and can create sub directories under /Data/Source/ in order to organize the templates as per convenience.

## Reports Location

The reports can be placed anywhere under directory /Data/Destination/. The report developer can create sub directories and give name to the output reports as per convenience.
