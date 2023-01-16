using System;
using System.IO;
using System.Runtime.InteropServices;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    public class AssemblyExamplesBase
    {
        static AssemblyExamplesBase()
        {
            MainDataDir = GetCodeBaseDir(System.Reflection.Assembly.GetExecutingAssembly());
            ArtifactsDir = new Uri(new Uri(MainDataDir), @"Data/Artifacts/").LocalPath;
            TemplatesDir = new Uri(new Uri(MainDataDir), @"Data/Templates/").LocalPath;
            DataSourcesDir = new Uri(new Uri(MainDataDir), @"Data/DataSources/").LocalPath;
            ImagesDir = new Uri(new Uri(MainDataDir), @"Data/Images/").LocalPath;
            LicenseDir = new Uri(new Uri(MainDataDir), @"Data/Licenses/").LocalPath;
        }

        [OneTimeSetUp]
        public static void OneTimeSetUp()
        {
            if (!Directory.Exists(ArtifactsDir))
                Directory.CreateDirectory(ArtifactsDir);

            SetUnlimitedLicense();
        }

        [SetUp]
        public static void SetUp()
        {
            Console.WriteLine($"Clr: {RuntimeInformation.FrameworkDescription}\n");
        }

        [OneTimeTearDown]
        public static void OneTimeTearDown()
        {
            if (Directory.Exists(ArtifactsDir))
                Directory.Delete(ArtifactsDir, true);
        }

        internal static void SetUnlimitedLicense()
        {
            // This is where the test license is on my development machine.
            string testLicenseFileName = Path.Combine(LicenseDir, "GroupDocs.Assembly for .NET.lic");

            if (File.Exists(testLicenseFileName))
            {
                License license = new License();
                license.SetLicense(testLicenseFileName);
            }
        }

        /// <summary>
        /// Returns the code-base directory.
        /// </summary>
        internal static string GetCodeBaseDir(System.Reflection.Assembly assembly)
        {
            Uri uri = new Uri(assembly.CodeBase);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("AssemblyExamples", StringComparison.Ordinal));
            
            return mainFolder;
        }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string MainDataDir { get; }

        /// <summary>
        /// Gets the path to the template documents.
        /// </summary>
        public static string TemplatesDir { get; }

        /// <summary>
        /// Gets the path of the data sources used by the code examples.
        /// </summary>
        internal static string DataSourcesDir { get; }

        /// <summary>
        /// Gets the path to the images used by the code examples.
        /// </summary>
        internal static string ImagesDir { get; }
        
        /// <summary>
        /// Gets the path to the artifacts produced by the code examples.
        /// </summary>
        internal static string ArtifactsDir { get; }

        /// <summary>
        /// Gets the path of the licenses.
        /// </summary>
        internal static string LicenseDir { get; }
    }
}