using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Npgsql;

namespace SMSSender
{
    class ConnectionManager
    {
        public String server = "localhost";
        public String port = "5432";
        public String user = "rhjasper";
        public String password = "rhjasper";
        public String database = "rhjasper";

        public String errorMsg = "";

        public NpgsqlConnection getConnection()
        {
            errorMsg = "";
            NpgsqlConnection conn;
            try
            {
                string strConnection = String.Format("Server={0};Port={1};" +
                    "User Id={2};Password={3};Database={4};",
                    server, port, user,
                    password, database);
                conn = new NpgsqlConnection(strConnection);
                conn.Open();
                return conn;
            }catch(Exception e){
                errorMsg = e.Message;
            }

            return null;
        }
    }
}
