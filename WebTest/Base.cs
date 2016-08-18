using System;
using System.Collections.Generic;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

namespace DAT_HHD
{
    /// <summary>
    /// Summary description for BaseData.
    /// </summary>
    public class BaseData : IBaseData
    {
        #region Variables
        private const string USP_EXCEPTIONLOG = "Usp_ExceptionLog";
        private string connectionString = String.Empty;
        public SqlConnection connection;
        protected SqlCommand sqlcmd;
        protected SqlTransaction sqltrans;
        public SqlConnection connectionException;
        #endregion

        public BaseData()
        {
            //
            // TODO: Add constructor logic here
            //
            connectionString = ConfigurationManager.ConnectionStrings["SAPCloudDBConnection"].ConnectionString;
            connection = new SqlConnection(connectionString);
            connectionException = new SqlConnection(connectionString);
        }
        #region IBaseData Members
        public void GetConnection()
        {
            connection.Open();
            sqlcmd = connection.CreateCommand();
        }

        public void CloseConnection()
        {
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();
            }
        }

        //converting ',",\r,\n,\t before inserting into the database
        public string convertto(string str)
        {
            str = str.Trim();
            str = str.Replace("'", "<sq>");
            str = str.Replace("\r", " ");
            str = str.Replace("\n", " ");
            str = str.Replace("\t", " ");
            str = str.Replace("\r\n", " ");
            str = str.Replace("\"", "<dq>");
            return (str.Trim());
        }

        //converting back ',",\r,\n,\t when retrieving from the database
        public string convertfrom(string str)
        {
            str = str.Trim();
            str = str.Replace("<sq>", "'");
            str = str.Replace("<dq>", "\"");
            return (str.Trim());
        }

        public SqlTransaction begintrans()
        {
            sqltrans = connection.BeginTransaction();
            return sqltrans;
        }

        public void commit()
        {
            sqltrans.Commit();
        }

        public string ExceptionLog(string error)
        {
            connectionException.Open();
            sqlcmd = connectionException.CreateCommand();
            this.sqlcmd.CommandText = USP_EXCEPTIONLOG;
            this.sqlcmd.CommandType = CommandType.StoredProcedure;
            this.sqlcmd.Parameters.Add("@Error", SqlDbType.VarChar).Value = error;
            this.sqlcmd.Parameters.Add("@Module", SqlDbType.VarChar).Value = "";

            try
            {
                this.sqlcmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                if (connectionException.State == ConnectionState.Open)
                {
                    connectionException.Close();
                }
            }
            return "Error In DataBase.Error Message : " + error + "Found";

        }

        public string ExceptionLog(string error, string module)
        {
            connectionException.Open();
            sqlcmd = connectionException.CreateCommand();
            this.sqlcmd.CommandText = USP_EXCEPTIONLOG;
            this.sqlcmd.CommandType = CommandType.StoredProcedure;
            this.sqlcmd.Parameters.Add("@Error", SqlDbType.VarChar).Value = error;
            this.sqlcmd.Parameters.Add("@Module", SqlDbType.VarChar).Value = module;

            try
            {
                this.sqlcmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                if (connectionException.State == ConnectionState.Open)
                {
                    connectionException.Close();
                }
            }
            return "Error In DataBase.Error Message : " + error + "Found";

        }

        public string Execute(string sql)
        {
            connectionException.Open();
            sqlcmd = connectionException.CreateCommand();
            this.sqlcmd.CommandText = sql;
            this.sqlcmd.CommandType = CommandType.Text;

            try
            {
                this.sqlcmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                if (connectionException.State == ConnectionState.Open)
                {
                    connectionException.Close();
                }
            }
            return "Success";

        }


        #endregion
    }
}
