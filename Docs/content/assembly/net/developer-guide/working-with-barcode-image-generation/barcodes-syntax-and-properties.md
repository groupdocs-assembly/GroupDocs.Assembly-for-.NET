---
id: barcodes-syntax-and-properties
url: assembly/net/barcodes-syntax-and-properties
title: Barcodes Syntax and Properties
weight: 1
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
{{< alert style="info" >}}This feature is only compatible with GroupDocs.Assembly for .NET 3.2.0 or later releases.{{< /alert >}}

<table class="sectionMacro" border="0" cellpadding="5" cellspacing="0" width="100%"><tbody><tr><td valign="top" width="50%"><div class="panel" style="border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;"><div class="panelHeader" style="border-bottom-width: 1px; background-color: rgb(176, 196, 222);"><b>Contents Summary</b></div><div class="panelContent"><style type="text/css">div.rbtoc1593026666860 { padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; }div.rbtoc1593026666860 ul { list-style-type: none; list-style-image: none; margin-left: 0px; }div.rbtoc1593026666860 li { margin-left: 0px; padding-left: 0px; }</style><div class="toc rbtoc1593026666860"><ul class="toc-indentation"><li><span class="TOCOutline">1</span> <a href="#BarcodesSyntaxandProperties-InsertingBarcodesDynamically">Inserting Barcodes Dynamically</a><ul class="toc-indentation"><li><span class="TOCOutline">1.1</span> <a href="#BarcodesSyntaxandProperties-ScalingBarcodeSymbols">Scaling Barcode Symbols</a></li><li><span class="TOCOutline">1.2</span> <a href="#BarcodesSyntaxandProperties-ConfiguringInitialScalingofBarcodeSymbols">Configuring Initial Scaling of Barcode Symbols</a></li><li><span class="TOCOutline">1.3</span> <a href="#BarcodesSyntaxandProperties-SettingHeightof1DBarcodeSymbols">Setting Height of 1D Barcode Symbols</a></li><li><span class="TOCOutline">1.4</span> <a href="#BarcodesSyntaxandProperties-CustomizingForeandBackColorsofBarcodeImages">Customizing Fore and Back Colors of Barcode Images</a></li><li><span class="TOCOutline">1.5</span> <a href="#BarcodesSyntaxandProperties-CustomizingFontandPlacementofBarcodeText">Customizing Font and Placement of Barcode Text</a></li><li><span class="TOCOutline">1.6</span> <a href="#BarcodesSyntaxandProperties-UsingBarcodeTextAuto-Correction">Using Barcode Text Auto-Correction</a></li><li><span class="TOCOutline">1.7</span> <a href="#BarcodesSyntaxandProperties-SettingBarcodeImageResolution">Setting Barcode Image Resolution</a></li><li><span class="TOCOutline">1.8</span> <a href="#BarcodesSyntaxandProperties-ExamplewithOutput">Example with Output</a></li></ul></li></ul></div></div></div></td><td valign="top">&nbsp;</td></tr></tbody></table>

## Inserting Barcodes Dynamically

You can insert barcode images to your documents dynamically using barcode tags. To declare a dynamically inserted barcode image within your template, perform the following steps:

1.  Add a textbox to your template at the place where you want a barcode image to be inserted.
2.  Set common image attributes such as size, rotation angle, and others for the textbox, making the textbox look like a barcode image without bars and text.
3.  Specify a barcode tag within the textbox using the following syntax:

```csharp
<<barcode [barcode_expression] -barcode_type>>
```

During run-time, an expression declared within a barcode tag is evaluated, its value is encoded using the specified barcode type, and the tag's containing textbox is filled with a corresponding barcode image. The tag itself is removed from the textbox then.  
To specify a barcode type within a barcode tag, you can use one of the predefined identifiers described in the following table:

| Identifier | Description |
| --- | --- |
| codabar | CODABAR Barcode |
| code11 | CODE 11 barcode |
| code39S | Standard CODE 39 barcode |
| code39E | Extended CODE 39 barcode |
| code93S | Standard CODE 93 barcode |
| code93E | Extended CODE 93 barcode |
| code128 | CODE 128 barcode |
| code128GS1 | GS1 CODE 128 barcode specification. The barcode text must contain parentheses for A |
| ean8 | EAN-8 barcode |
| ean13 | EAN-13 barcode |
| ean14 | EAN-14 barcode |
| scc14 | SCC-14 barcode |
| sscc18 | SSCC-18 barcode |
| upca | UPC-A barcode |
| upce | UPC-E barcode |
| isbn | ISBN barcode |
| issn | ISSN barcode |
| ismn | ISMN barcode |
| stf | Standard 2 of 5 barcode |
| itf | Interleaved 2 of 5 barcode |
| mtf | Matrix 2 of 5 barcode |
| ip25 | Italian Post 25 barcode |
| iatatf | IATA 2 of 5 barcode. IATA (International Air Transport Assosiation) uses this barcode for the management of air cargo |
| itf14 | ITF14 barcode |
| itf6 | ITF-6 barcode |
| msi | MSI Plessey barcode |
| vin | VIN (Vehicle Identification Number) barcode |
| dpi | Deutschen Post barcode, also known as Identcode, CodeIdentcode, German Postal 2 of 5 Identcode, Deutsche Post AG Identcode, Deutsche Frachtpost Identcode, Deutsche Post AG (DHL) |
| dpl | Deutsche Post Leitcode barcode, also known as German Postal 2 of 5 Leitcode, CodeLeitcode, Leitcode, Deutsche Post AG (DHL) |
| opc | OPC (Optical Product Code) barcode, also known as VCA barcode, VCA OPC, Vision Council of America OPC barcode |
| pzn | PZN barcode, also known as Pharma Zentral Nummer, Pharmazentralnummer |
| code16K | Code 16K barcode |
| pharmacode | Represents Pharmacode barcode, also known as Code32 |
| dm | DataMatrix barcode |
| qr | QR Code barcode |
| aztec | Aztec barcode |
| pdf417 | Pdf417 barcode |
| macroPdf417 | MacroPdf417 barcode |
| dmGS1 | DataMatrix barcode with GS1 string format |
| microPdf417 | MicroPdf417 barcode |
| qrGS1 | QR barcode with GS1 string format |
| maxiCode | MaxiCode barcode |
| dotCode | DotCode barcode |
| ap | Australia Post Customer barcode |
| postnet | Postnet barcode |
| planet | Planet barcode |
| oneCode | USPS OneCode barcode |
| rm4scc | RM4SCC barcode. RM4SCC (Royal Mail 4-state Customer Code) is used for automated mail sort process in UK |
| databarOD | Databar omni-directional barcode |
| databarT | Databar truncated barcode |
| databarL | Databar limited barcode |
| databarE | Databar expanded barcode |
| databarES | Databar expanded stacked barcode |
| databarS | Databar stacked barcode |
| databarSOD | Databar stacked omni-directional barcode |
| sp | Singapore Post barcode |
| ape | Australian Post Domestic eParcel barcode |
| spp | Swiss Post Parcel barcode. Supported types: Domestic Mail, International Mail, Additional Services |
| patchCode | Patch code barcode |
| code32 | Code32 barcode |
| dltf | DataLogic 2 of 5 barcode |
| dkix | Dutch KIX barcode |
| codablockF | Codablock F barcode |
| codablockFGS1 | GS1 Codablock F barcode |

### Scaling Barcode Symbols

You can scale barcode symbols of dynamically inserted barcode images specifying the scale attribute for a barcode tag as follows.

```csharp
<<barcode [barcode_expression] -barcode_type scale="scaling_hint">>
```

A scale attribute affects the width and height of a 2D barcode symbol and the width of a 1D barcode symbol. The value of a scale attribute represents the percent of barcode symbol scaling.  
The value of a scale attribute, however, is just a hint. That is, actual scaling of a barcode symbol can be not precisely equal to the specified one. The reason is that scaling can damage a barcode symbol in a way it to become unreadable by barcode scanners. That is why, the engine calculates actual scaling not to bring such a damage trying to be as close to the specified scaling as possible at the same time.

### Configuring Initial Scaling of Barcode Symbols

You can affect initial scaling of all barcode symbols generated by a single DocumentAssembler instance through the **DocumentAssembler.BarcodeSettings.BaseXDimension** and **DocumentAssembler.BarcodeSettings.BaseYDimension** properties. These properties define the smallest width and height of the unit of barcode modules respectively. That is, for example, to make all barcode symbols twice larger by default, you can use the following code:

```csharp
DocumentAssembler assembler = new DocumentAssembler();
assembler.BarcodeSettings.BaseXDimension *= 2f;
assembler.BarcodeSettings.BaseYDimension *= 2f;

```

Values of the **DocumentAssembler.BarcodeSettings.BaseXDimension** and **DocumentAssembler.BarcodeSettings.BaseYDimension** properties are measured in units defined by the **DocumentAssembler.BarcodeSettings.GraphicsUnit** property.

{{< alert style="warning" >}}When the scale attribute is provided for a barcode tag, actual x- and y-dimensions are calculated by the engine as the corresponding percent of base x- and y-dimensions.{{< /alert >}}

### Setting Height of 1D Barcode Symbols

The height of a 1D barcode symbol is not affected by the scale attribute of the corresponding barcode tag. However, you can set the height applying the barHeight attribute as follows:

```csharp
<<barcode [barcode_expression] -barcode_type barHeight="height">>
```

The value of a barHeight attribute specifies the percent of the overall barcode image height. That is, for example, to set the height of a barcode symbol equal to two thirds of its overall barcode image height, you can use the following barcode tag:

```csharp
<<barcode ["736192"] -codabar barHeight="67">>
```

### Customizing Fore and Back Colors of Barcode Images

The fore color of a dynamically inserted barcode image is used to render the barcode's symbol and text. This color is derived from the font of the corresponding barcode tag. The background color of the image is derived from the solid fill of the textbox containing the tag.

### Customizing Font and Placement of Barcode Text

The font used to render barcode text of a dynamically inserted barcode image derives the following attributes from the font of the corresponding barcode tag:

*   Name
*   Size
*   Style (bold and italic)
*   Color

The placement of barcode text of a dynamically inserted barcode image – leftward, rightward, or centered – is defined by text alignment of the corresponding barcode tag.

### Using Barcode Text Auto-Correction

For a dynamically inserted barcode image, the engine automatically corrects barcode text that does not conform the barcode's specification, if possible.

{{< alert style="warning" >}}The auto-correction is not possible for barcodes of some types like Databar.{{< /alert >}}

To disable the auto-correction making the engine throw an exception in case of an invalid barcode text, you can set the **DocumentAssembler.BarcodeSettings.UseAutoCorrection** property to false.

### Setting Barcode Image Resolution

You can set horizontal and vertical resolution for dynamically inserted barcode images in DPI (dots per inch) using the **DocumentAssembler.BarcodeSettings.XResolution** and **DocumentAssembler.BarcodeSettings.YResolution** properties respectively. That is, for example, to set both horizontal and vertical resolution to 300 DPI, you can use the following code:

```csharp
DocumentAssembler assembler = new DocumentAssembler();
assembler.BarcodeSettings.XResolution = 300;
assembler.BarcodeSettings.YResolution = 300;

```

{{< alert style="warning" >}}By default, horizontal and vertical resolution of dynamically inserted barcode images is 96 DPI{{< /alert >}}

### Example with Output

```csharp
<<barcode ["736192"] -codabar>>
```

Generated bar code for the above mentioned syntax is:

![](https://github.com/groupdocs-assembly/GroupDocs.Assembly-for-.NET/blob/master/Examples/Data/Screenshots/Barcode.png?raw=true)
