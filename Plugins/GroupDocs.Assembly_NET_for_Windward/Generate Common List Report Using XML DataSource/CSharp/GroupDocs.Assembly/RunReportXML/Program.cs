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
    //ExStart:Generate common list in assembly-Csharp
    public class Program
    {
        public const string customerXMLfile = "../../../../../Data/Data Source/Customers.xml";
        static void Main(string[] args)
        {
            //Setting up source open presentation template
            FileStream template = File.OpenRead("../../../../../Data/Samples/Source/Common List.docx");
            //Setting up destination open presentation report 
            FileStream output = File.Create("../../../../../Data/Samples/Destination/Xml Common List.docx");

            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate Bulleted List Report in open presentation format
                assembler.AssembleDocument(template, output, GetCustomListData(), "ds");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static DataSet GetCustomListData()
        {
            try
            {
                DataSet mainDs = new DataSet();
                System.IO.FileStream fsReadXml = new System.IO.FileStream(customerXMLfile, System.IO.FileMode.Open);
                mainDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                mainDs.Tables[0].TableName = "Customers";
                return mainDs;
            }
            catch
            {
                return null;
            }
        }
    }
    //ExEnd:Generate common list in assembly-Csharp
}
