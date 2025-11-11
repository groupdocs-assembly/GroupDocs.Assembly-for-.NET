using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    public static class MultipleDataSources
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # MultipleDataSources : Assemble document using multiple data sources (object, JSON, XML) \n");

            string[] templates = { "Multiple data source.odt", "Multiple data source.ods", "Multiple data source.odp" };

            foreach (string template in templates)
            {
                string extension = Path.GetExtension(template);

                DocumentAssembler assembler = new DocumentAssembler();
                
                assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, template),
                    Path.Combine(Constants.OutputPath, "MultipleDataSources" + extension),
                    new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"),
                    new DataSourceInfo(DataLayer.GetProductsDataJson(), "products"));
            }
        }
    }
}


