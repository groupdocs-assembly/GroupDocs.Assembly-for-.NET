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
    //ExStart:Generate numbered list in assembly-Csharp
    public class Program
    {
        public const string productXMLfile = "../../../../../Data/Data Source/Products.xml";
        static void Main(string[] args)
        {
            //Setting up source open presentation template
            FileStream template = File.OpenRead("../../../../../Data/Samples/Source/Numbered List_XML.docx");
            //Setting up destination open presentation report 
            FileStream output = File.Create("../../../../../Data/Samples/Destination/Xml Numbered List.docx");

            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate Bulleted List Report in open presentation format
                assembler.AssembleDocument(template, output, GetProductsData(), "product");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static DataSet GetProductsData()
        {
            try
            {
                DataSet mainDs = new DataSet();
                System.IO.FileStream fsReadXml = new System.IO.FileStream(productXMLfile, System.IO.FileMode.Open);
                mainDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                mainDs.Tables[0].TableName = "ds";
                return mainDs;
            }
            catch
            {
                return null;
            }
        }
    }
    //ExEnd:Generate numbered list in assembly-Csharp
}
