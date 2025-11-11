using System;
using System.IO;

namespace GroupDocs.Assembly.Examples.CSharp
{
    public static class Constants
    {
        private static readonly string BasePath = FindBasePath();

        private static string GetBaseDirectory()
        {
            // Use AppContext.BaseDirectory for .NET Core/.NET 6.0 (available in .NET Framework 4.6+)
            // Fallback to AppDomain.CurrentDomain.BaseDirectory for older .NET Framework
            try
            {
                return AppContext.BaseDirectory ?? AppDomain.CurrentDomain.BaseDirectory;
            }
            catch
            {
                return AppDomain.CurrentDomain.BaseDirectory;
            }
        }

        private static string FindBasePath()
        {
            string baseDirectory = GetBaseDirectory();
            string currentPath = Path.GetFullPath(baseDirectory);

            // Search up the directory tree to find the folder containing the "Data" directory
            // This works regardless of the project type (Framework, Core, or .NET 6.0)
            while (currentPath != null && !string.IsNullOrEmpty(currentPath))
            {
                string dataPath = Path.Combine(currentPath, "Data");
                string templatesPath = Path.Combine(dataPath, "Templates");

                // Check if Data\Templates exists (marker for the correct base path)
                if (Directory.Exists(templatesPath))
                {
                    return currentPath;
                }

                // Move up one directory level
                string parentPath = Path.GetDirectoryName(currentPath);
                
                // Stop if we've reached the root or can't go higher
                if (parentPath == null || parentPath == currentPath)
                {
                    break;
                }
                
                currentPath = parentPath;
            }

            // Fallback: try going up 3 levels from base directory (original logic)
            return Path.GetFullPath(Path.Combine(baseDirectory, "..", "..", ".."));
        }

        public static readonly string TemplatesPath = Path.Combine(BasePath, "Data", "Templates");
        public static readonly string DataSourcesPath = Path.Combine(BasePath, "Data", "DataSources");
        public static readonly string ImagesPath = Path.Combine(BasePath, "Data", "Images");
        public static readonly string OutputPath = Path.Combine(BasePath, "Data", "Artifacts");

        public static readonly string TemplatesDir = EnsureTrailingSeparator(TemplatesPath);
        public static readonly string DataSourcesDir = EnsureTrailingSeparator(DataSourcesPath);
        public static readonly string ImagesDir = EnsureTrailingSeparator(ImagesPath);

        private static string EnsureTrailingSeparator(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return path;
            }

            return path.EndsWith(Path.DirectorySeparatorChar.ToString()) ? path : path + Path.DirectorySeparatorChar;
        }

        static Constants()
        {
            Directory.CreateDirectory(Constants.OutputPath);
        }
    }
}
