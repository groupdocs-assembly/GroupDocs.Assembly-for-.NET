---
id: evaluation-limitations-and-licensing
url: assembly/net/evaluation-limitations-and-licensing
title: Evaluation Limitations and Licensing
weight: 5
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}You can use GroupDocs.Assembly without the license. The usage and functionalities are pretty much same as the licensed one but you will face few limitations while using the non-licensed API.{{< /alert >}}

## Evaluation Version Limitations

You can easily download GroupDocs.Assembly for evaluation. The evaluation download is the same as the purchased download. The evaluation version simply becomes licensed when you add a few lines of code to apply the license. You will face following limitations while using the API without the license.  

| Document | Spreadsheet | Presentation |
| --- | --- | --- |
| GroupDocs.Assembly (without a license specified) generates the document reports with full product functionality, but it inserts an evaluation watermark at the top of the document | When a Spreadsheet report is generated it contains an extra worksheet with Evaluation Copyright Warning on it, you can't hide the Evaluation Copyright Warning (the extra worksheet) | When a presentation report is generated using GroupDocs.Assembly, an Evaluation Watermark is injected at the center of each slide |
| It limits the maximum document size to several hundred paragraphs | When running your program, you can only generate 100 Spreadsheet reports using GroupDocs.Assembly. If your application exceeds this number, an exception will be thrown |  |
|  | You can't re-specify the script path (e.g in the web.config file) |  |

## Licensing 

The license file contains details such as the product name, number of developers it is licensed to, subscription expiry date and so on. It contains the digital signature, so don't modify the file. Even inadvertent addition of an extra line break into the file will invalidate it. You need to set a license before utilizing GroupDocs.Assembly API if you want to avoid its evaluation limitations. 

The license can be loaded from a file or stream object. The easiest way to set a license is to put the license file in the same folder as the GroupDocs.Assembly.dll file and specify the file name, without a path, as shown in the examples below.

### Applying License from File

The code below will explain how to call the license file.
{{< gist GroupDocsGists db63054e45146654ff4efeda01f556a2 >}}

### Applying License from Stream

The following example shows how to load a license from a stream.
{{< gist GroupDocsGists e6e23199ab4c88f7f60f191a65cdfdc6 >}}

{{< alert style="info" >}}

Calling License.SetLicense multiple times is not harmful but simply wastes processor time. If you are developing a Windows Forms or console application, call License.SetLicense in your startup code, before using GroupDocs.Assembly classes.

When developing an ASP.NET application, you can call License.SetLicense from the Global.asax.cs (Global.asax.vb) file in the Application_Start protected method. It is called once when the application starts. Do not call License.SetLicense from within Page_Load methods since it means the license will be loaded every time a web page is loaded.

{{< /alert >}}

{{< alert style="info" >}}See this video about applying the license.{{< /alert >}}

### Applying Metered License

{{< alert style="info" >}}You can also set Metered license as an alternative to license file. It is a new licensing mechanism that will be used along with existing licensing method. It is useful for the customers who want to be billed based on the usage of the API features. For more details, please refer to Metered Licensing FAQ section.{{< /alert >}}

Here are the simple steps to use the `Metered` class.

1.  Create an instance of `Metered` class.
2.  Pass public & private keys to `SetMeteredKey` method.
3.  Do processing (perform task).
4.  call method `GetConsumptionQuantity` of the `Metered` class.
5.  It will return the amount/quantity of API requests that you have consumed so far.
6.  call method `GetConsumptionCredit` of the `Metered` class (Since version 19.7).
7.  It will return the credit that you have consumed so far.

Following is the sample code demonstrating how to use `Metered` class.
{{< gist GroupDocsGists 7d4158b476ad2d413f1e04b03ffc1794 >}}