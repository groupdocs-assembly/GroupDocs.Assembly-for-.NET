using System.IO;
using AssemblyExamples.Data;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class BarCodes : AssemblyExamplesBase
    {
        [TestCase("Barcode.docx")]
        [TestCase("Barcode.xlsx")]
        [TestCase("Barcode.pptx")]
        public void AddBarcode(string template)
        {
            string extension = Path.GetExtension(template);

            //ExStart:AddBarcode
            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + template,
                ArtifactsDir + "BarCodes.AddBarcode" + extension,
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            //ExEnd:AddBarcode
        }

        /// <summary>
        /// Changing the resolution of barcode images while saving the document.
        /// Uniform resolution in DPI across both the X and Y axes is supported by version 20.8 or greater.
        /// </summary>
        [TestCase(72, "Barcode.docx")]
        [TestCase(1600, "Barcode.docx")]
        public void BarcodeResolution(float resolution, string sourceTemplateFilename)
        {
            //ExStart:BarcodeResolution
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.BarcodeSettings.Resolution = resolution;

            assembler.AssembleDocument(TemplatesDir + sourceTemplateFilename,
                ArtifactsDir + $"BarCodes.BarcodeResolution.{resolution}dpi.docx",
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            //ExEnd:BarcodeResolution
        }

        /// <summary>
        /// Changing the scaling of the barcode image within its containing shape while saving the document.
        /// </summary>
        [Test]
        public void BarcodeScale()
        {
            //ExStart:SetBarcodeScale
            DocumentAssembler assembler = new DocumentAssembler();

            assembler.BarcodeSettings.BaseXDimension *= 0.5f;
            assembler.BarcodeSettings.BaseYDimension *= 0.5f;

            assembler.AssembleDocument(TemplatesDir + "Barcode.docx",
                ArtifactsDir + "BarCodes.BarcodeScale.docx",
                new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            //ExEnd:SetBarcodeScale
        }
    }
}
