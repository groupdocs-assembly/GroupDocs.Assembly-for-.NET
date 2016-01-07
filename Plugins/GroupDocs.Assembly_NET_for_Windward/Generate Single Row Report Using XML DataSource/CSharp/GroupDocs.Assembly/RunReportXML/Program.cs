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
    //ExStart:Generate single row in assembly-Csharp
    public class Program
    {
        public const string customerXMLfile = "../../../../../Data/Data Source/Customers.xml";
        static void Main(string[] args)
        {
            //Setting up source open presentation template
            FileStream template = File.OpenRead("../../../../../Data/Samples/Source/Single Row.docx");
            //Setting up destination open presentation report 
            FileStream output = File.Create("../../../../../Data/Samples/Destination/Xml Single Row.docx");

            try
            {
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate Bulleted List Report in open presentation format
                assembler.AssembleDocument(template, output, GetSingleRow(), "customer");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static DataRow GetSingleRow()
        {
            try
            {
                DataSet singleCustomer = new DataSet();
                FileStream readProductXML = new FileStream(customerXMLfile, FileMode.Open);
                singleCustomer.ReadXml(readProductXML, XmlReadMode.ReadSchema);
                singleCustomer.Tables[0].TableName = "Customers";
                return singleCustomer.Tables["Customers"].Rows[0];
            }
            catch
            {
                return null;
            }
        }
    }
    //ExEnd:Generate single row in assembly-Csharp
}
