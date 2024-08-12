using System;
using System.Data.SqlClient;

namespace SyncFaturasSageEmail.Connects
{
    class Connect_HT_server
    {
        private string DBhost, DBname, DBuser, DBpass;
        private SqlConnection Connection;
        public static string ConnectionString = null;

        public Connect_HT_server(string host, string dbname, string user, string pass)
        {
            DBhost = host;
            DBname = dbname;
            DBuser = user;
            DBpass = pass;

            ConnectionString = "Data Source='" + DBhost + "'; Initial Catalog='" + DBname + "'; User Id='" + DBuser + "'; Password='" + DBpass + "'; Trusted_Connection=False";
            Connection = new SqlConnection();
            Connection.ConnectionString = ConnectionString;
        }
        public void ConnInit()
        {
            Connection.Close();
            Connection.Open();
        }
        public void ConnEnd()
        {
            Connection.Close();
        }
        public SqlConnection Conn { get { return Connection; } }
    }
}