using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.Sql;

namespace useSQL
{
    public sealed class useSQL : WQT_Welders.SQLDataBaseConnection
    {
        System.Data.Common.DbCommand newcommand = new SqlCommand();

        //Properties
        private string _storedProcedure;
        public override string StoredProcedure
        {
            get { return _storedProcedure; }
            set
            {
                _storedProcedure = value;

                newcommand.CommandType = System.Data.CommandType.StoredProcedure;
                newcommand.CommandText = value;
                newcommand.Parameters.Clear();
            }
        }

        private System.Data.Common.DbConnection _dbConnection;
        public override System.Data.Common.DbConnection dbconnection
        {
            get { return _dbConnection; }
        }

        static string _connectionString;
        public override string conString
        {
            get
            {
                return _connectionString;
            }
            set
            {
                base.conString = value;
                base.paramscollection = newcommand.Parameters;

                _connectionString = value;

                _dbConnection = new SqlConnection(value);
                if (_dbConnection.State == System.Data.ConnectionState.Closed)
                {
                    _dbConnection.ConnectionString = value;
                }
            }
        }

        //Constructor   
        public useSQL()
            : this("")
        {
        }
        public useSQL(string connectionString)
            : base(connectionString)
        {
            conString = connectionString;

            if (_dbConnection.State == System.Data.ConnectionState.Closed)
                _dbConnection.ConnectionString = _connectionString;
        }

        //Methods
        public override void Dispose()
        {
            newcommand.Dispose();
            _dbConnection.Close();
        }
        public override bool isDataBaseConnected()
        {
            try
            {
                if (dbconnection.State == System.Data.ConnectionState.Closed)
                    dbconnection.Open();
                return true;
            }
            catch (Exception e)
            {
                return false;
                throw e;
            }
        }
        public override System.Data.DataTable PerformCommand()
        {
            System.Data.DataTable dtable = new System.Data.DataTable();
            try
            {
                using (System.Data.Common.DbConnection con = new SqlConnection(conString))
                {
                    con.Open();
                    newcommand.Connection = con;
                    dtable.Load(newcommand.ExecuteReader());

                    return dtable;
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                dbconnection.Dispose();
                newcommand.Dispose();
            }
        }
        public override System.Data.DataTable PerformCommand(string storedProcedure, System.Data.Common.DbParameter[] parameter)
        {
            System.Data.DataTable dtable = new System.Data.DataTable();

            try
            {
                if (isDataBaseConnected())
                {
                    newcommand.CommandType = System.Data.CommandType.StoredProcedure;
                    newcommand.CommandText = StoredProcedure;
                    newcommand.Parameters.Clear();

                    for (int intIndex = 0; intIndex <= parameter.Length - 1; intIndex++)
                        newcommand.Parameters.Add(parameter[intIndex]);

                    if (dbconnection.State == System.Data.ConnectionState.Closed)
                        dbconnection.Open();

                    newcommand.Connection = dbconnection;
                    dtable.Load(newcommand.ExecuteReader());
                }
                else
                {
                    throw new Exception("ExecuteCommand(string, DbParameter): Connection not established.");
                }

                return dtable;

            }
            catch (SqlException sqlex)
            {
                throw new Exception(sqlex.Message);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (dbconnection != null)
                    dbconnection.Close();

                dbconnection.Dispose();
                newcommand.Dispose();
            }

        }
        public override System.Data.DataTable PerformQuery(string sqlQuery)
        {
            System.Data.DataTable dtable = new System.Data.DataTable();

            try
            {

                if (isDataBaseConnected())
                {
                    newcommand.CommandTimeout = 0;
                    newcommand.CommandType = System.Data.CommandType.Text;
                    newcommand.CommandText = sqlQuery;

                    newcommand.Parameters.Clear();
                    newcommand.Connection = dbconnection;
                    dtable.Load(newcommand.ExecuteReader());
                }
                else
                {
                    throw new Exception("ExecuteQuery(string): Connection not established.");
                }
                return dtable;
            }
            catch (SqlException sqlex)
            {
                throw new Exception(sqlex.Message);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                if (dbconnection != null)
                    dbconnection.Close();

                dbconnection.Dispose();
                newcommand.Dispose();
            }
        }
        public override System.Data.DataTable PerformQuery(string sqlQuery, System.Data.Common.DbParameter[] parameter)
        {
            System.Data.DataTable dtable = new System.Data.DataTable();

            try
            {
                if (isDataBaseConnected())
                {
                    newcommand.CommandType = System.Data.CommandType.Text;
                    newcommand.CommandText = sqlQuery;

                    newcommand.Parameters.Clear();
                    for (int i = 0; i <= parameter.Length - 1; i++)
                        newcommand.Parameters.Add(parameter[i]);

                    if (dbconnection.State == System.Data.ConnectionState.Closed)
                        dbconnection.Open();

                    dtable.Load(newcommand.ExecuteReader());
                }
                else
                {
                    throw new Exception("ExecuteQuery(String, DbParameter): Connection not established.");
                }

                return dtable;
            }
            catch (System.Data.Common.DbException sqlex)
            {
                throw new Exception(sqlex.Message);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (dbconnection != null)
                    dbconnection.Close();

                dbconnection.Dispose();
                newcommand.Dispose();
            }

        }


        //Events
        public override event WQT_Welders.SQLDataBaseConnection.connectionEventHandler connectionServerStateEvent;
        protected override void OnconnectionStateChange(WQT_Welders.ConnectionServerStateEventArgs e)
        {
            if (connectionServerStateEvent != null)
                connectionServerStateEvent(this, e);
        }
    }
}