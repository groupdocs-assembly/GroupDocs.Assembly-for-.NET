using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.AssemblyExamples.BusinessLayer;
using GroupDocs.Assembly;

namespace GroupDocs.AssemblyExamples
{
    public class Program
    {
        static void Main(string[] args)
        {
            //CommonUtilities.ProductLicense();

            string folderName = "WordTemplates";
            string fileName = "Common List.docx";
            string dataSourceName = "customers";
            string updatedFileName = CommonUtilities.ChangeFileName(fileName);

            try
            {
                DocumentAssembler assembler = new DocumentAssembler();
                assembler.AssembleDocument(CommonUtilities.DocumentSourceFolderPath(fileName, folderName), CommonUtilities.DocumentOutputFolderPath(updatedFileName, folderName), DataLayer.PopulateData(), dataSourceName);
            }
            catch (Exception ex)
            {
                
                Console.WriteLine(ex.Message);
            }
           
            //If dataSourceName is products call ProductsData data source 
            //If dataSourceName is orders call OrdersData data source 
            //If dataSourceName is customers call PopulateData data source
            //If dataSourceName is customer call CustomerData data source
        }
    }
}

