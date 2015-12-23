using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GroupDocs.ReportGenerator.DataAccessLayer;
using System.Data;

namespace GroupDocs.ReportGenerator.BusinessLayer
{
   /// <summary>
   /// Enum to check the type of DBObject
   /// </summary>
    public enum DBObjects
    {
        Tables,
        Views
    }
    /// <summary>
    /// Business class of DB datasource 
    /// </summary>
    public class BusinessDB : ReportGenerator
    {

        private string connectionString;

        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; }
        }

       /// <summary>
       /// Constructor which accepts connection String
       /// </summary>
       /// <param name="dbConnectionString"></param>
        public BusinessDB(string dbConnectionString)
        {
            connectionString = dbConnectionString;
        }

        public BusinessDB()
        {
        }
        /// <summary>
        /// To check the state and validity of connection
        /// </summary>
        /// <returns></returns>
        public bool IsValidConnection()
        {
            try
            {
                DBHandler dbHandler = new DBHandler(connectionString);
                dbHandler.openConnection();
                if (dbHandler.getConnection().State == ConnectionState.Open)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }
        /// <summary>
        /// Get the names of DBO e.g. Tables or Views
        /// </summary>
        /// <param name="dbo"></param>
        /// <returns></returns>
        public DataSet getDBObjectNames(DBObjects dbo)
        {
            DBHandler dbHandler = new DBHandler(connectionString);
            String Query = "";
            switch (dbo)
            {
                case DBObjects.Tables:
                    Query = "SELECT name FROM sys.tables order by name";
                    break;
                case DBObjects.Views:
                    Query = "SELECT name FROM sys.views order by name";
                    break;
            }
            DataSet ds = dbHandler.ExecuteSelectQuery(Query);

            return ds;

        }
        /// <summary>
        ///  Generates Report if just dbo are selected
        /// </summary>
        /// <param name="objName"></param>
        /// <param name="destinationPath"></param>
        /// <param name="templateSourcePath"></param>
        public void GenerateReport_DBO(String objName, String destinationPath, String templateSourcePath)
        {
            DBHandler dbHandler = new DBHandler(connectionString);

            String Query = "SELECT * FROM " + objName;

            DataSet ds = dbHandler.ExecuteSelectQuery(Query);
            TemplateSourcePath = templateSourcePath;
            ReportDestinationPath = destinationPath;
            ds.Tables[0].TableName = objName;
            GenerateReportDB(ds);
        }
        /// <summary>
        /// Generates the report if custom query is selected.
        /// </summary>
        /// <param name="Query"></param>
        /// <param name="destinationPath"></param>
        /// <param name="templateSourcePath"></param>
        public void GenerateReport_Query(String Query, String destinationPath, String templateSourcePath)
        {
            DBHandler dbHandler = new DBHandler(connectionString);

            DataSet ds = dbHandler.ExecuteSelectQuery(Query);
            TemplateSourcePath = templateSourcePath;
            ReportDestinationPath = destinationPath;
            GenerateReportDB(ds);
        }
    }
}
