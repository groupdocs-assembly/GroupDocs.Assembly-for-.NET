using GroupDocs.Assembly;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupDocs.AssemblyExamples.BusinessLayer
{
    public class CommonUtilities
    {
        //ExStart:CommonUtilities

        public const string sourceFolderPath = "../../../../Data/SourceFolder/";
        public const string destinationFolderPath = "../../../../Data/DestinationFolder/";
        public const string licensePath = "../../GroupDocs.Assembly Product Family.lic";

        #region DocumentDirectories

       

        //ExStart:DocumentDirectories

        /// <summary>
        /// Takes file name and source folder name as arguments. 
        /// </summary>
        /// <param name="sourceFileName">Source file name</param>
        /// <param name="sourceFolder">Source file's folder name</param>
        /// <returns>Returns explicit path by combining source folder path, source folder name and file name.</returns>

        public static string DocumentSourceFolderPath(string sourceFileName, string sourceFolder)
        {
            return Path.Combine(Path.GetFullPath(sourceFolderPath), sourceFolder, sourceFileName);
        }

        /// <summary>
        /// Takes file name and destination folder name as arguments. 
        /// </summary>
        /// <param name="outputFileName">output file name</param>
        /// <param name="destinationFolder">output file's folder name</param>
        /// <returns>Returns explicit path by combining destination folder path, destination folder name and output file name.</returns>

        public static string DocumentOutputFolderPath(string outputFileName, string destinationFolder)
        {
            return Path.Combine(Path.GetFullPath(destinationFolderPath), destinationFolder, outputFileName);
        }

        //ExEnd:DocumentDirectories
        #endregion


        #region ProductLicense

        /// <summary>
        /// Set product's license
        /// </summary>

        public static void ProductLicense()
        {
            License lic = new License();
            lic.SetLicense(licensePath);
        }

        #endregion

        #region ChangeFileName

        /// <summary>
        /// Takes source file name as argument and append it with datetime and "CS".
        /// </summary>
        /// <param name="sourceFileName">Source file name</param>
        /// <returns>Appended file name</returns>
        
        public static string ChangeFileName(string sourceFileName)
        {
            string currentDateTime = DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss");
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(sourceFileName);
            string extension = Path.GetExtension(sourceFileName);
            string alterFileName = nameWithoutExtension + "-" + currentDateTime + "-CS";
            string updatedFileName = alterFileName + extension;

            return updatedFileName;
        }

        #endregion

        //ExEnd:CommonUtilities
    }
}
