using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    public static class TableReports
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # TableReports : Working with table reports \n");

            DataTable();
            CellsMerging();
        }

        public static void DataTable()
        {
            string[] templates = { "Data table.docx", "Data table.xlsx", "Data table.pptx", "Data table.msg" };

            foreach (string template in templates)
            {
                string extension = Path.GetExtension(template);

                DocumentAssembler assembler = new DocumentAssembler();

                assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, template),
                    Path.Combine(Constants.OutputPath, "TableReports.DataTable" + extension),
                    new DataSourceInfo(DataLayer.ExcelData(), "ds"));
            }
        }

        public static void CellsMerging()
        {
            string[] templates = { "Merging table cells dynamically.docx", "Merging table cells dynamically.xlsx", "Merging table cells dynamically.pptx" };

            foreach (string template in templates)
            {
                string extension = Path.GetExtension(template);

                DocumentAssembler assembler = new DocumentAssembler();
                
                assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, template),
                    Path.Combine(Constants.OutputPath, $"TableReports.CellsMerging_{extension}.pdf"),
                    new LoadSaveOptions(FileFormat.Pdf),
                    new DataSourceInfo(DataLayer.GetCustomerData(), "customer"));
            }
        }
    }
}

