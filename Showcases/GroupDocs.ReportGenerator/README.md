## Web Report Generator Powered by GroupDocs.Assembly 

Web Report Generator is an ASP.NET Web Forms Application that demonstrates the core functionalities provided by GroupDocs.Assembly for .NET.
The Web Report Generator allows user to easily create reports based on:

+ **Predefined custom templates** 
+ **Your own database (MS SQL Server local OR Remote) tables, views and also with custom query**

The UI of the application is kept very simple and developers can enhance the features as their own needs since it is an
Open Source Project.

## Prerequisite

+ Visual Studio 2012 or later.
+ Any latest browsers (Firefox, Chrome ,IE9 or greater, Safari etc. )
+ GroupDocs.Assembly 17.9.

## Features 

+ Connect to MS SQL Server local OR remote by providing connection string.
+ Populate and use Tables from connected database.
+ Populate and use Views from connected database.
+ Create and use custom query over connected database.
+ Upload and use MS Word Template file.
+ Upload and use MS Excel Template file.
+ Upload and use MS Power Point Template file.
+ Generate Report in a format as per source Template Document format.
+ Download generated report from web browser.

### Add Missing APIs

Follow the given steps to rebuild the solution which would automatically add missing references:
* Extract the downloaded project and open the solution file in Visual Studio
* Right click on solution and press "Enable NuGet package Restore"
* Rebuild the project

In other case, it is possible that Visual Studio is unable to automatically add these third party APIs references due to Visual Studio version differences. In this case, please add references of the following APIs manually:

1) [GroupDocs.Assembly](https://www.nuget.org/packages/GroupDocs.Assembly/)
