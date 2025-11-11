using System;
using System.IO;

namespace GroupDocs.Assembly.Examples.CSharp.AdvancedUsage.Infrastructure
{
    public abstract class ExampleBase
    {
        protected static string TemplatesDir => ExampleResourcePaths.TemplatesDir;
        protected static string DataSourcesDir => ExampleResourcePaths.DataSourcesDir;
        protected static string ImagesDir => ExampleResourcePaths.ImagesDir;
        protected static string ArtifactsDir { get; } = EnsureArtifactsDir();

        private static string EnsureArtifactsDir()
        {
            string path = Path.Combine(Constants.OutputPath, "Artifacts");
            Directory.CreateDirectory(path);
            return path + Path.DirectorySeparatorChar;
        }

        protected static void EnsureOutputDirectory(string relativePath)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return;
            }

            string fullPath = Path.Combine(Constants.OutputPath, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Constants.OutputPath);
        }

        protected static string CombineTemplate(string fileName)
        {
            return Path.Combine(TemplatesDir, fileName);
        }

        protected static string CombineDataSource(string fileName)
        {
            return Path.Combine(DataSourcesDir, fileName);
        }

        protected static string CombineImage(string fileName)
        {
            return Path.Combine(ImagesDir, fileName);
        }

        protected static void RequireTemplate(string fileName)
        {
            if (!File.Exists(CombineTemplate(fileName)))
            {
                throw new FileNotFoundException($"Template '{fileName}' was not found at '{TemplatesDir}'.", CombineTemplate(fileName));
            }
        }

        protected static void LogMissingTemplate(string fileName)
        {
            Helper.WriteError($"Template '{fileName}' was not found. Skipping example.");
        }

        protected static void LogMissingData(string fileName)
        {
            Helper.WriteError($"Data source '{fileName}' was not found. Skipping example.");
        }
    }
}


