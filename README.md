# .NET Document Automation & Reporting API

An on-premise [Document Automation Engine](https://products.groupdocs.com/assembly/net) that accepts a template document and data to generate documents as per syntax defined by the template. Data can be merged from various sources including CSV, XML, OData, JSON, .NET Objects & more.

<p align="center">

  <a title="Download GroupDocs.Assembly for .NET Sources" href="https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

Directory | Description
--------- | -----------
[Examples](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/tree/master/Examples)  | Contain package of all .NET examples and sample files that will help you learn how to use GroupDocs.Assembly for NET to its full extent. 
[Plugins](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/tree/master/Plugins)  | Analyze the progress while generating reports, compare code and template preparation. 
[Showcases](https://github.com/groupdocsassembly/GroupDocs_Assembly_NET/tree/master/Showcases)  | Web Report Generator is an ASP.NET Web Forms Application that demonstrates the core functionalities provided by GroupDocs.Assembly for .NET.

## Report Generation via .NET

- Wide range of supported [document formats](https://docs.groupdocs.com/display/assemblynet/Supported+Document+Formats).
- Document generation using data from [various sources](https://docs.groupdocs.com/assembly/net/working-with-multiple-data-sources/).
- LINQ-based template syntax.
- Manipulate data using formulae & [sequential data](https://docs.groupdocs.com/display/assemblynet/Template+Syntax+-+Part+2+of+2#TemplateSyntax-Part2of2-OutputtingSequentialData) operations.
- Supports Upper, Lower, Capital, First-Cap content formatting.
- Apply Ordinal, Cardinal, Alphabetic & Numeric formatting to template syntax.
- Apply attributes to email message body and dynamically add attachments.
- [Generate barcode labels](https://docs.groupdocs.com/display/assemblynet/Working+with+Barcode+Image+Generation) in reports.
- Apply conditional formatting & formatting based on type such as numeric, text, image & date-time.
- Supports NEXT field analogue of Microsoft Word.
- Update fields during Word document assembly.
- Apply & calculate formula during Excel file assembly.
- Generate various report types, such as, Charts, Lists, Tables etc.
- Load templates from HTML as well as save assembled documents as HTML with resources.

## Read & Write Microsoft Office Formats

**Microsoft Word:** DOC, DOT, DOCX, DOCM, DOTX, DOTM, RTF\
**Microsoft Excel:** XLSX, XLSM, XLTX, XLTM, XLSB, XLS, XLT\
**Microsoft PowerPoint:** PPTX, PPTM, PPSX, PPSM, POT, POTX, POTM, PPT, PPS

## Other Supported Formats

**OpenOffice:** ODS, ODT, OTT, OTP, ODP, OXPS\
**Email:** EML, MSG, EMLX\
**Fixed Layout:** PDF, XPS\
**Web:** HTML, MHTML\
**Images:** TIFF, SVG\
**Other:** XML, XAML, TXT, EPUB, PS, PCL, MD

## Platform Independence

GroupDocs.Assembly  for .NET can be used to build applications for Windows, Mac OS X as well as for Linux. Developers may code in PHP, VBScript, Delphi & C++ programming languages while using GroupDocs.Assembly for .BET via COM Interop.

## Getting Started with GroupDocs.Assembly for .NET

Are you ready to give GroupDocs.Assembly for .NET a try? Simply execute `Install-Package GroupDocs.Assembly` from Package Manager Console in Visual Studio to fetch & reference GroupDocs.Assembly assembly in your project. If you already have GroupDocs.Assembly for .Net and want to upgrade it, please execute `Update-Package GroupDocs.Assembly` to get the latest version.

## Generate DOCX from Markdown (MD) Template via C# Code

```csharp
//Setting up source document template
const String strDocumentTemplate = "Markdown Templates/ReadMe.md";
//Setting up destination Markdown reports
const String strDocumentReport = "Word Reports/ReadMe Out.docx";
//Setting up description variable
const string description = "GroupDocs.Assembly for .NET is a class library that enables you to generate documents in popular " +
    "office and email file formats based upon template documents and data obtained from various sources " +
    "including databases, XML, JSON, OData, objects of custom .NET types, external documents, and more.";

DocumentAssembler assembler = new DocumentAssembler();
//Assemble Document
assembler.AssembleDocument(
CommonUtilities.GetSourceDocument(strDocumentTemplate),
CommonUtilities.SetDestinationDocument(strDocumentReport),
new DataSourceInfo("GroupDocs.Assembly for .NET", "product"),
new DataSourceInfo(description, "description"));
}
```

## Use Word Table as a Source to Assemble a Document

```csharp
// assemble a document using the external document table as a data source.
DocumentAssembler assembler = new DocumentAssembler();
assembler.AssembleDocument(CommonUtilities.GetSourceDocument(strDocumentTemplate),
                           CommonUtilities.SetDestinationDocument(strDocumentReport),
                           new DataSourceInfo(DataLayer.ExcelData(), "contracts"));
```

[Home](https://www.groupdocs.com/) | [Product Page](https://products.groupdocs.com/assembly/net) | [Documentation](https://docs.groupdocs.com/assembly/net/) | [Demo](https://products.groupdocs.app/assembly/family) | [API Reference](https://apireference.groupdocs.com/assembly/net) | [Examples](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET) | [Blog](https://blog.groupdocs.com/category/assembly/) | [Search](https://search.groupdocs.com/) | [Free Support](https://forum.groupdocs.com/c/assembly) | [Temporary License](https://purchase.groupdocs.com/temp-license/100170)
