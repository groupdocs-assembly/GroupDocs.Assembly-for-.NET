using GroupDocs.Assembly;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RunReportXML
{
    public class Program
    {
        public const string productXMLfile = "../../../Samples/Products.xml";
        static void Main(string[] args)
        {
            //ExStart:GenerateBulletedListFromXMLinOpenPresentationFormat
            //Setting up source open presentation template
            FileStream template = File.OpenRead("../../../Samples/Bulleted List_XML.docx");
            //Setting up destination open presentation report 
            FileStream output = File.Create("../../../Samples/Xml Bulleted List.docx");

            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate Bulleted List Report in open presentation format
                assembler.AssembleDocument(template, output, GetAllDataFromXML(), "ds");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //ExEnd:GenerateBulletedListFromXMLinOpenPresentationFormat
        }
        public static DataSet GetAllDataFromXML()
        {
            try
            {
                DataSet mainDs = new DataSet();
                System.IO.FileStream fsReadXml = new System.IO.FileStream(productXMLfile, System.IO.FileMode.Open);
                mainDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                mainDs.Tables[0].TableName = "Product";
                return mainDs;
            }
            catch
            {
                return null;
            }
        }
    }
}
