using System;
using System.Data;
using System.IO;
using System.Text;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    /// <summary>
    /// Shows how to call a custom static function, registered through
    /// <see cref="DocumentAssembler.KnownTypes"/>, that takes a single <see cref="System.Data.DataRow"/>
    /// parameter. When a JSON data source is used, an object node such as <c>root</c> is passed to the
    /// function as one <see cref="DataRow"/>, so a signature like <c>FormatName(DataRow)</c> is resolved
    /// and invoked — no need to declare the parameter as an <c>IEnumerable</c>/collection.
    /// </summary>
    public static class AssembleUsingCustomDataRowFunction
    {
        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # AssembleUsingCustomDataRowFunction : Call a KnownTypes function that takes a single DataRow from a JSON object node \n");

            // A single JSON object exposed under the "root" node.
            const string json = @"{
  ""root"": {
    ""title"": ""Ms"",
    ""firstName"": ""Jane"",
    ""lastName"": ""Doe"",
    ""manager"": { ""firstName"": ""John"", ""lastName"": ""Smith"" }
  }
}";

            // `root` is passed to the custom function as a single DataRow.
            const string sourceString = "Contact: <<[ContactFunctions.FormatName(root)]>>";

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.KnownTypes.Add(typeof(ContactFunctions));

            JsonDataLoadOptions options = new JsonDataLoadOptions { AlwaysGenerateRootObject = true };

            byte[] sourceBytes = Encoding.UTF8.GetBytes(sourceString);
            byte[] targetBytes;

            using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
            using (MemoryStream targetStream = new MemoryStream())
            {
                JsonDataSource dataSource = new JsonDataSource(jsonStream, options);
                assembler.AssembleDocument(sourceStream, targetStream,
                    new DataSourceInfo(dataSource));
                targetBytes = targetStream.ToArray();
            }

            string outputPath = Path.Combine(Constants.OutputPath, "AssembleUsingCustomDataRowFunction.txt");
            File.WriteAllBytes(outputPath, targetBytes);
        }
    }

    /// <summary>
    /// A custom function type registered via <see cref="DocumentAssembler.KnownTypes"/>. Must be a
    /// top-level, publicly visible type so the template engine can resolve it by name.
    /// </summary>
    public static class ContactFunctions
    {
        // Accepts a single DataRow (a JSON object node such as `root`) and reads its columns.
        public static string FormatName(DataRow contact)
        {
            if (contact == null)
            {
                return string.Empty;
            }

            return string.Format("{0} {1} {2}", contact["title"], contact["firstName"], contact["lastName"]);
        }
    }
}
