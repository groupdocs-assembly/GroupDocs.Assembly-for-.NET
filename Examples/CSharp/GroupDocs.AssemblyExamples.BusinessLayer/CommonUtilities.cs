using GroupDocs.Assembly;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;

namespace GroupDocs.AssemblyExamples.BusinessLayer
{
    //ExStart:CommonUtilities
    public static class CommonUtilities
    {

        public const string sourceFolderPath = "../../../../Data/Source/";
        public const string dataFolderPath = "../../../../Data/";
        public const string destinationFolderPath = "../../../../Data/Destination/";
        public const string dataSourcesFolderPath = "../../../../Data/Data Sources/";
        //ExStart:LicenseFilePath
        public const string licensePath = "D:/GroupDocs.Total.NET.lic";
        //ExEnd:LicenseFilePath

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
        /// Takes source folder name as argument. 
        /// </summary>
        public static string GetSourceFolder(string sourceFolder)
        {
            return Path.Combine(Path.GetFullPath(dataFolderPath), sourceFolder);
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

        /// <summary>
        /// Takes source file name as argument. 
        /// </summary>
        /// <param name="sourceFileName">Source file name</param>
        /// <returns>Returns explicit path by combining data source folder path and source file name.</returns>
        public static string GetDataSourceDocument(string sourceFileName)
        {
            return Path.Combine(Path.GetFullPath(dataSourcesFolderPath), sourceFileName);
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

        //ExStart:metered licensing 
        /// <summary>
        /// provide metered licensing
        /// </summary>
        public static void MeteredLicensing()
        {

            // Set metered license public and private keys.
            Metered metered = new Metered();
            metered.SetMeteredKey("PublicKey", "PrivateKey");

            //Do something here....

            // and get consumption quantity
            decimal consumptionQuantitiy = GroupDocs.Assembly.Metered.GetConsumptionQuantity();

            // get consumption credit (Supported since version 19.7)
            decimal consumptionCredit = GroupDocs.Assembly.Metered.GetConsumptionCredit();
        }
        //ExEnd:metered licensing 
        #endregion

        #region ToADOTable
        //ExStart:ConvertToDataTable
        /// <summary>
        /// It takes delegate and varlist IEnumberable
        /// </summary>
        /// <typeparam name="T">Template</typeparam>
        /// <param name="varlist">IEnumerable varlist</param>
        /// <param name="fn">Delegate as parameter</param>
        /// <returns>It returns DataTable</returns>
        public static DataTable ToADOTable<T>(this IEnumerable<T> varlist, ConvertDataTable.CreateRowDelegate<T> fn)
        {
            DataTable dtReturn = new DataTable();
            PropertyInfo[] oProps = null;
            foreach (T rec in varlist)
            {
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    foreach (PropertyInfo pi in oProps)
                    {
                        Type colType = pi.PropertyType; if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }
                        dtReturn.Columns.Add(new DataColumn(pi.Name, colType));
                    }
                }
                DataRow dr = dtReturn.NewRow(); foreach (PropertyInfo pi in oProps)
                {
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue(rec, null);
                }
                dtReturn.Rows.Add(dr);
            }
            return (dtReturn);
        }
        #endregion
        public static class ConvertDataTable
        {
            public delegate object[] CreateRowDelegate<T>(T t);
        }
        //ExEnd:ConvertToDataTable

        public static class FileUtil
        {
            public static string GetBytesAsBase64(string path)
            {
                return Convert.ToBase64String(File.ReadAllBytes(path));
            }
        }
    }

    //ExEnd:CommonUtilities
}
