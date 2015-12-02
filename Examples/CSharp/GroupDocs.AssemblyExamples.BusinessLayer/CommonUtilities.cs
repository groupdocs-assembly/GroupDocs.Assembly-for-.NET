using GroupDocs.Assembly;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupDocs.AssemblyExamples.BusinessLayer
{
    //ExStart:CommonUtilities
    public class CommonUtilities
    {

        public const string sourceFolderPath = "../../../../Data/Source/";
        public const string destinationFolderPath = "../../../../Data/Destination/";
        public const string licensePath = "../../GroupDocs.Assembly Product Family.lic";

        #region DocumentDirectories

        //ExStart:DocumentDirectories
        /// <summary>
        /// Takes source file name as argument. 
        /// </summary>
        /// <param name="sourceFileName">Source file name</param>
        /// <returns>Returns explicit path by combining source folder path and source file name.</returns>
        public static string GetSourceDocument(string sourceFileName)
        {
            return Path.Combine(Path.GetFullPath(sourceFolderPath), sourceFileName);
        }
        /// <summary>
        /// Takes output file name as argument. 
        /// </summary>
        /// <param name="outputFileName">output file name</param>
        /// <returns>Returns explicit path by combining destination folder path and output file name.</returns>

        public static string SetDestinationDocument(string outputFileName)
        {
            return Path.Combine(Path.GetFullPath(destinationFolderPath), outputFileName);
        }

        //ExEnd:DocumentDirectories
        #endregion

        #region ProductLicense
        //ExStart:ApplyLicense
        /// <summary>
        /// Set product's license
        /// </summary>
        public static void ApplyLicense()
        {
            License lic = new License();
            lic.SetLicense(licensePath);
        }
        //ExEnd:ApplyLicense
        #endregion
    }
    //ExEnd:CommonUtilities
}
