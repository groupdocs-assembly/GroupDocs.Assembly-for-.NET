using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class Conversions : AssemblyExamplesBase
    {
        /// <summary>
        /// Change target file format.
        /// Features is supported by version 18.9 or greater.
        /// </summary>
        [Test]
        public void ChangeTargetFileFormat()
        {
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Bubble chart.docx",
                ArtifactsDir + "Conversions.ChangeTargetFileFormat.pdf",
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));

            assembler.AssembleDocument(TemplatesDir + "Bubble chart.docx",
                ArtifactsDir + "Conversions.ChangeTargetFileFormatUsingExplicitSpecifying.pdf",
                new LoadSaveOptions(FileFormat.Pdf),
                new DataSourceInfo(DataLayer.GetOrdersData(), "orders"));
        }

        /// <summary>
        /// Saving template POT presentation document to assembled Presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        [Test]
        public void PotToPptx()
        {
            //ExStart:PotToPptx
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Conversion template.pot",
                ArtifactsDir + "Conversions.PotToPptx.pptx",
                new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
            //ExEnd:PotToPptx
        }

        /// <summary>
        /// Saving template OTP presentation document to assembled Presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        [Test]
        public void OtpToPptx()
        {
            //ExStart:OtpToPptx
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Conversion template.otp",
                ArtifactsDir + "Conversions.OtpToPptx.pptx",
                new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
            //ExEnd:OtpToPptx
        }

        /// <summary>
        /// Saving assembled Presentation document to template OTP presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        [Test]
        public void PptxToOtp()
        {
            //ExStart:PptxToOtp
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Conversion template.pptx",
                ArtifactsDir + "Conversions.PptxToOtp.otp",
                new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
            //ExEnd:PptxToOtp
        }

        /// <summary>
        /// Saving assembled Presentation document to template OTP presentation document.
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        [Test]
        public void PptxToPot()
        {
            //ExStart:PptxToPot
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Conversion template.pptx",
                ArtifactsDir + "Conversions.PptxToPot.pot",
                new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
            //ExEnd:PptxToPot
        }

        /// <summary>
        /// Saving assembled Presentation document to template POT presentation document (stream).
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        [Test]
        public void PptxToPotAsStream()
        {
            //ExStart:PptxToPotAsStream
            using (Stream templateStream =
                new FileStream(TemplatesDir + "Conversion template.pptx", FileMode.Open))
            {
                using (Stream resultPotStream =
                    new FileStream(ArtifactsDir + "Conversions.PptxToPotAsStream.pot", FileMode.CreateNew))
                {
                    DocumentAssembler assembler = new DocumentAssembler();

                    assembler.AssembleDocument(templateStream, resultPotStream, new LoadSaveOptions(FileFormat.Pot),
                        new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
                }
            }
            //ExEnd:PptxToPotAsStream
        }

        /// <summary>
        /// Saving assembled Presentation document to template OTP presentation document (stream).
        /// Features is supported by version 20.6 or greater.
        /// </summary>
        [Test]
        public void PptxToOtpAsStream()
        {
            //ExStart:PptxToOtpAsStream
            using (Stream templateStream =
                new FileStream(TemplatesDir + "Conversion template.pptx", FileMode.Open))
            {
                using (Stream resultPotStream =
                    new FileStream(ArtifactsDir + "Conversions.PptxToOtpAsStream.otp", FileMode.CreateNew))
                {
                    DocumentAssembler assembler = new DocumentAssembler();

                    assembler.AssembleDocument(templateStream, resultPotStream, new LoadSaveOptions(FileFormat.Otp),
                        new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
                }
            }
            //ExEnd:PptxToOtpAsStream
        }

        /// <summary>
        /// Load an XLT spreadsheet and save it to the XLSX format.
        /// Feature is supported by version 20.12 or greater.
        /// </summary>
        [Test]
        public void XltToXlsx()
        {
            //ExStart:XltToXlsx
            using (Stream templateStream =
                new FileStream(TemplatesDir + "Conversion template.xlt", FileMode.Open))
            {
                using (Stream resultPotStream = new FileStream(ArtifactsDir + "Conversions.XltToXlsx.xlsx", FileMode.Create))
                {
                    DocumentAssembler assembler = new DocumentAssembler();

                    assembler.AssembleDocument(templateStream, resultPotStream, new LoadSaveOptions(FileFormat.Xlsx),
                        new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
                }
            }
            //ExEnd:XltToXlsx
        }

        /// <summary>
        /// Load an XLSX spreadsheet and save it to the XLT format.
        /// Feature is supported by version 20.12 or greater.
        /// </summary>
        [Test]
        public void XlsxToXlt()
        {
            //ExStart:XlsxToXlt
            using (Stream templateStream =
                new FileStream(TemplatesDir + "Conversion template.xlsx", FileMode.Open))
            {
                using (Stream resultPotStream =
                    new FileStream(ArtifactsDir + "XlsxToXlt.xlt", FileMode.Create))
                {
                    DocumentAssembler assembler = new DocumentAssembler();

                    assembler.AssembleDocument(templateStream, resultPotStream, new LoadSaveOptions(FileFormat.Xlt),
                        new DataSourceInfo("GroupDocs.Assembly for .NET", "product"));
                }
            }
            //ExEnd:XlsxToXlt
        }
    }
}
