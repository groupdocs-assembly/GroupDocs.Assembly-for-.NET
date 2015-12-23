using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace GroupDocs.ReportGenerator.DataAccessLayer
{
   /// <summary>
   /// General Database Handler
   /// </summary>
    public class DBHandler
    {
        SqlConnection sqlCon = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        string _ConnectionString = "";

        public string ConnectionString
        {
            get { return _ConnectionString; }
            set { _ConnectionString = value; }
        }

        public DBHandler(String ConnString)
        {
            _ConnectionString = ConnString;
        }
        public SqlConnection getConnection()
        {
            return sqlCon;
        }
        /// <summary>
        /// Method to open a Database connection
        /// </summary>
        /// <returns></returns>
        public SqlConnection openConnection()
        {

            if (sqlCon.State != ConnectionState.Open)
            {

                sqlCon.ConnectionString = _ConnectionString;
                sqlCon.Open();
            }
            return sqlCon;
        }
        /// <summary>
        /// Method to close a connection
        /// </summary>
        public void closeConnection()
        {
            if (sqlCon.State != ConnectionState.Closed)
            {
                sqlCon.Close();
            }
        }

        /// <summary>
        /// If wanna execute Insert,Update and Delete
        /// </summary>
        /// <param name="sqlQuery">Query in the form of String</param>
        /// <returns>0 or 1</returns>
        public int Execute(string sqlQuery)
        {
            try
            {

                SqlCommand cmd = new SqlCommand(sqlQuery, sqlCon);
                openConnection();
                int i = cmd.ExecuteNonQuery();
                closeConnection();
                return i;
            }
            catch (Exception ex)
            {
                closeConnection();
                throw ex;
            }

        }
        /// <summary>
        /// To exectute select statement
        /// </summary>
        /// <param name="sql">Query in the form of String</param>
        /// <returns></returns>
        public DataSet ExecuteSelectQuery(string sql)
        {

            DataSet myDs = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter(sql, sqlCon);
            openConnection();
            adapter.Fill(myDs);
            closeConnection();
            return myDs;
        }
    }

}
