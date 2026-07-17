using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using GroupDocs.Assembly;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage
{
    /// <summary>
    /// Shows that GroupDocs.Assembly resolves members a data object inherits from its <b>base
    /// interfaces</b>. Here the Boolean flag <c>Disabled</c> is declared on the base interface
    /// <see cref="IToggleable"/>, while the collection elements are referenced through the most-derived
    /// interface <see cref="IFeatureItem"/>. A template expression that applies the logical-NOT operator
    /// to that inherited flag (<c>!item.Disabled</c>) is therefore evaluated as a Boolean, so a filter
    /// such as <c>Items.Where(i =&gt; !i.Disabled)</c> works exactly like it does for plain classes.
    /// </summary>
    public static class AssembleUsingInterfaceDataModel
    {
        // The Boolean flag lives on the deepest base interface ...
        public interface IToggleable
        {
            bool Disabled { get; }
        }

        // ... the display title on the middle one ...
        public interface INamedItem : IToggleable
        {
            string Title { get; set; }
        }

        // ... and the elements are exposed through the most-derived interface.
        public interface IFeatureItem : INamedItem
        {
            string Category { get; set; }
        }

        public class FeatureItem : IFeatureItem
        {
            public string Title { get; set; }
            public bool Disabled { get; set; }
            public string Category { get; set; }
        }

        public class FeatureList
        {
            public string Name { get; set; }

            // The collection element type is the interface IFeatureItem, so the engine resolves
            // `Disabled` against an interface hierarchy rather than a concrete class.
            public ICollection<IFeatureItem> Items { get; set; } = new List<IFeatureItem>();
        }

        public static void Run()
        {
            Console.WriteLine("\n--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("[Example Advanced Usage] # AssembleUsingInterfaceDataModel : Filter an interface-typed collection using a Boolean flag declared on a base interface \n");

            FeatureList featureList = new FeatureList
            {
                Name = "Release 2026",
                Items = new List<IFeatureItem>
                {
                    new FeatureItem { Title = "Legacy import", Disabled = true,  Category = "Import" },
                    new FeatureItem { Title = "Cloud sync",    Disabled = false, Category = "Sync" },
                    new FeatureItem { Title = "Dark theme",     /* Disabled defaults to false */ Category = "UI" },
                }
            };

            // `!item.Disabled` reads a Boolean property inherited from the IToggleable base interface.
            const string sourceString =
                "Enabled features in <<[Name]>>:\r\n" +
                "<<foreach [item in Items.Where(i => !i.Disabled)]>>- <<[item.Title]>> (<<[item.Category]>>)\r\n<</foreach>>";

            DocumentAssembler assembler = new DocumentAssembler();

            byte[] sourceBytes = Encoding.UTF8.GetBytes(sourceString);
            byte[] targetBytes;

            using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
            using (MemoryStream targetStream = new MemoryStream())
            {
                assembler.AssembleDocument(sourceStream, targetStream,
                    new DataSourceInfo(featureList, "features"));
                targetBytes = targetStream.ToArray();
            }

            string outputPath = Path.Combine(Constants.OutputPath, "AssembleUsingInterfaceDataModel.txt");
            File.WriteAllBytes(outputPath, targetBytes);
        }
    }
}
