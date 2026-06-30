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
                string? envLicense = Environment.GetEnvironmentVariable(Constants.ENV_LICENSE_FILE);

                if (File.Exists(Constants.LicensePath))
                {
                    License license = new License();
                    license.SetLicense(Constants.LicensePath);
                    Console.WriteLine("License set successfully.");
                }
                else if (!string.IsNullOrEmpty(envLicense) && File.Exists(envLicense))
                {
                    License license = new License();
                    license.SetLicense(envLicense);
                    Console.WriteLine("License set successfully.");
                }
                else
                {
                    Console.WriteLine("\nWe do not ship any license with this example. " +
                                      "\nVisit the GroupDocs site to obtain either a temporary or permanent license. " +
                                      "\nLearn more about licensing at https://purchase.groupdocs.com/faqs/licensing. " +
                                      "\nLear how to request temporary license at https://purchase.groupdocs.com/temporary-license.");
                }
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


