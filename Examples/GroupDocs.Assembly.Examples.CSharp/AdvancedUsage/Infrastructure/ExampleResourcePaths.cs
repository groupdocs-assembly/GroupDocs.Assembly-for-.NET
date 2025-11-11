using System.IO;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage.Infrastructure
{
    internal static class ExampleResourcePaths
    {
        static ExampleResourcePaths()
        {
        }

        public static string TemplatesDir => Constants.TemplatesDir;
        public static string DataSourcesDir => Constants.DataSourcesDir;
        public static string ImagesDir => Constants.ImagesDir;
        public static string ArtifactsDir => Path.Combine(Constants.OutputPath, "Artifacts") + Path.DirectorySeparatorChar;
    }
}


