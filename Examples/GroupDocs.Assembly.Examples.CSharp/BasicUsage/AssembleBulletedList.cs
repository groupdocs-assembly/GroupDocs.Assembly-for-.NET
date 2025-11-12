using System;
using System.IO;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Examples.CSharp.Data;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleBulletedList
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleBulletedList : Assemble document with bulleted list from data source \n");

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(Path.Combine(Constants.TemplatesPath, "Bulleted list.odp"),
                Path.Combine(Constants.OutputPath, "AssembleBulletedList.odp"),
                new DataSourceInfo(DataLayer.GetAllDataFromXml(), "ds"));
        }
    }
}

