using GroupDocs.Assembly;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GroupDocs.AssemblyExamples.BusinessLayer
{
    //ExStart:CommonUtilities
    public static class CommonUtilities
    {

        public const string sourceFolderPath = "../../../../Data/Source/";
        public const string destinationFolderPath = "../../../../Data/Destination/";
        //ExStart:LicenseFilePath
        public const string licensePath = "D:/GroupDocs.Total.lic";
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
    }
   
    //ExEnd:CommonUtilities
}
