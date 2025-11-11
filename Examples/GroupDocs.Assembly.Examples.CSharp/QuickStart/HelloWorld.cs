using System;
using System.IO;

namespace GroupDocs.Assembly.Examples.CSharp.QuickStart
{
    using GroupDocs.Assembly;

    public static class HelloWorld
    {
        public class HelloWorldData
        {
            public string Name { get; set; }
            public DateTime Now { get; set; }
        }

        public static void Run()
        {
            // Use a simple string template and an object as data source
            string templatePath = Path.Combine(Constants.OutputPath, "template.txt");
            File.WriteAllText(templatePath, "Hello, {{Name}}! Today is {{Now:dd MMM yyyy}}.");

            var data = new HelloWorldData { Name = "World", Now = DateTime.Now };
            string output = Path.Combine(Constants.OutputPath, "hello-world.txt");

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(templatePath, output, new DataSourceInfo(data, "data"));
        }
    }
}


