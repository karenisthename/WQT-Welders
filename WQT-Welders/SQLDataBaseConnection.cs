using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.Sql;

namespace WQT_Welders
{
    public abstract class SQLDataBaseConnection : IDisposable
    {
          //Properties
        public abstract string StoredProcedure
        {
            get;
            set;
        }
        public abstract System.Data.Common.DbConnection dbconnection
        {
            get;
        }
        
        static string _connectionString;
        public virtual string conString
        {
            get{ return _connectionString;}
            set { _connectionString = value; }
        }

        System.Data.Common.DbParameterCollection _parameterCollection;
        public virtual System.Data.Common.DbParameterCollection paramscollection
        {
            get { return _parameterCollection; }
            set { _parameterCollection = value; }
        }

        //Constructor
        public SQLDataBaseConnection()
            : this("")
        { }
        public SQLDataBaseConnection(string connectionString)
        {
            _connectionString = connectionString;
        }

        //Methods
        public abstract void Dispose();
        public abstract bool isDataBaseConnected();
        public abstract System.Data.DataTable PerformCommand();
        public abstract System.Data.DataTable PerformCommand(string storedProcedure, System.Data.Common.DbParameter[] parameter);
        public abstract System.Data.DataTable PerformQuery(string sqlQuery);
        public abstract System.Data.DataTable PerformQuery(string sqlQuery, System.Data.Common.DbParameter[] parameter);

        //Events
        public delegate void connectionEventHandler(object sender, ConnectionServerStateEventArgs e);
        public virtual event connectionEventHandler connectionServerStateEvent;
        protected virtual void OnconnectionStateChange(ConnectionServerStateEventArgs e)
        {
            if (connectionServerStateEvent != null)
            {
                connectionServerStateEvent(this, e);
                e.statusMessage = "Connection Status is OPEN.";
            }
            else
            {
                e.statusMessage = "Connection Status is CLOSED.";
            }
        }
    }

    public sealed class ConnectionServerStateEventArgs : EventArgs
    {
        ConnectionServerStateEnum statusConnection;
        private string connectionMessage;

        //constructor
        public ConnectionServerStateEventArgs()
        { }
        public ConnectionServerStateEventArgs(System.Data.Common.DbConnection connection)
        {
            if (connection != null)
            {
                if (connection.State == System.Data.ConnectionState.Open)
                {
                    connectionMessage = "Connection Status: OPEN";
                    statusConnection = ConnectionServerStateEnum.Open;
                }
                else if (connection.State == System.Data.ConnectionState.Closed)
                {
                    try
                    {
                        connection.Open();
                        statusConnection = ConnectionServerStateEnum.Open;
                        connectionMessage = "Connection Status Changed to: CLOSED";
                    }
                    catch (Exception e)
                    {
                        connection.Close();
                        statusConnection = ConnectionServerStateEnum.Closed;
                        throw new Exception(e.Message);
                    }
                }
                else
                {
                    statusConnection = ConnectionServerStateEnum.Unknown;
                    connectionMessage = "Connection Status: UNKNOWN";
                }
            }
            else {
                connection.Close();
                statusConnection = ConnectionServerStateEnum.Closed;
                connectionMessage = "Connection Status: CLOSED.";
            }
        }

        public string statusMessage
        {
            get { return connectionMessage; }
            set{ connectionMessage = value; }
        }

        public ConnectionServerStateEnum ConnectionServerState
        {
            get { return statusConnection; }
        }
    }
    public enum ConnectionServerStateEnum
    {
        Closed=0,
        Open=1,
        Unknown= 2
    }
}
