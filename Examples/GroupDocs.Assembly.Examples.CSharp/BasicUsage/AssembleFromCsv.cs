using System;
using System.IO;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    using GroupDocs.Assembly.Data;

    public static class AssembleFromCsv
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleFromCsv : Assemble document from CSV data source \n");

            string template = Path.Combine(Constants.TemplatesPath, "Data destination with nested elements.txt");
            string csvPersons = Path.Combine(Constants.DataSourcesPath, "Persons.csv");
            string output = Path.Combine(Constants.OutputPath, "AssembleFromCsv.docx");
                        
            var dataSource = new CsvDataSource(csvPersons, new CsvDataLoadOptions(true));

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(template, output,
                new DataSourceInfo(dataSource, "persons"));
        }
    }
}


