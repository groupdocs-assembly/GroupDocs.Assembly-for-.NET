using System;
using System.IO;
using GroupDocs.Assembly;

namespace GroupDocs.Assembly.Examples.CSharp.QuickStart
{
    public static class SetLicenseFromFile
    {
        public static void Run()
        {
            string licenseFileName = "GroupDocs.Assembly.lic";
            // uncomment if license in Environment LIC_PATH
            // licenseFileName = Path.Combine(Environment.GetEnvironmentVariable("LIC_PATH"), licenseFileName);

            try
            {
                License license = new License();
                license.SetLicense(licenseFileName);
                Console.WriteLine($"License set successfully from: {licenseFileName}");
            }
            catch (Exception ex)
            {
                Helper.WriteError($"License was not set: {ex.Message}");
                // License file not found, display message
                Console.WriteLine();
                Console.WriteLine("We do not ship license in this example.");
                Console.WriteLine("To get temporary licenses go to:");
                Console.WriteLine("https://purchase.groupdocs.com/temp-license/100170");
                Console.WriteLine();
            }
        }
    }
}


