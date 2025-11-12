using System;
using System.IO;
using System.Text;
using GroupDocs.Assembly;

namespace GroupDocs.Assembly.Examples.CSharp.BasicUsage
{
    public static class AssembleFromStringTemplate
    {
        public class StringTemplateData
        {
            public string Name { get; set; }
            public DateTime Date { get; set; }
        }

        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Basic Usage] # AssembleFromStringTemplate : Assemble document from string template \n");

            DocumentAssembler assembler = new DocumentAssembler();

            const string sourceString = @"Hello, <<[Name]>>! Today is <<[Date]:""dd MMMM yyyy"">>.";
            byte[] sourceBytes = Encoding.UTF8.GetBytes(sourceString);
            byte[] targetBytes;

            using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
            {
                using (MemoryStream targetStream = new MemoryStream())
                {
                    var data = new StringTemplateData { Name = "World", Date = DateTime.Now };
                    assembler.AssembleDocument(sourceStream, targetStream,
                        new DataSourceInfo(data, "data"));
                    targetBytes = targetStream.ToArray();
                }
            }

            string outputPath = Path.Combine(Constants.OutputPath, "AssembleFromStringTemplate.txt");
            File.WriteAllBytes(outputPath, targetBytes);
        }
    }
}

