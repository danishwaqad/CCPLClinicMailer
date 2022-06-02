using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCPLClinicMailer.DAL
{
    class DB_Helper
    {
        SqlConnection Connection;
        SqlCommand Command;
        SqlDataAdapter Adapter;
        public string DatabaseName { get; set; }
        public string CompanyID { get; set; }

        SqlConnectionStringBuilder Builder;

        public static string ClassConnectionString = "";

        public DB_Helper(bool UseProjectConnection = false)
        {
            if (UseProjectConnection)
                Connection = new SqlConnection(@ClassConnectionString);
            else
                //Connection = new SqlConnection(ConfigurationManager.ConnectionStrings["GPMailer"].ConnectionString);
                ConnectionBuilder();
        }

        public void ConnectionBuilder()
        {
            try
            {
                Builder = new SqlConnectionStringBuilder();
                if (DatabaseName != null)
                {
                    Builder.InitialCatalog = DatabaseName;
                }
                else
                {
                    Builder.InitialCatalog = ConfigurationManager.AppSettings["dbName"];
                }
                Builder.DataSource = ConfigurationManager.AppSettings["Server"];
                Builder.UserID = ConfigurationManager.AppSettings["dbID"];
                Builder.Password = ConfigurationManager.AppSettings["dbPass"];
                Connection = new SqlConnection();
                Builder.ConnectTimeout = 100;
                Connection.ConnectionString = Builder.ConnectionString;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //public DBHelper(string ConnectionString)
        //{
        //    Connection = new SqlConnection(@ConnectionString);
        //}

        public void OpenSQLConnection()
        {
            ConnectionBuilder();
            if (Connection.State == System.Data.ConnectionState.Closed)
                Connection.Open();
        }
        public SqlConnection GetSQLConnection()
        {
            return Connection;
        }
        public void CloseSQLConnection()
        {
            if (Connection.State == System.Data.ConnectionState.Open)
                Connection.Close();
        }
        //public long GetFirstFreeId(string TableName, string colName = "Id")
        //{
        //    long Id = 0;
        //    long PreId = 0;

        //    string query = "SELECT " + colName + " FROM " + TableName + " ORDER BY " + colName + " ASC";
        //    Command = new SqlCommand(query, GetSQLConnection());
        //    OpenSQLConnection();
        //    SqlDataReader Reader = Command.ExecuteReader();
        //    while (Reader.Read())
        //    {
        //        Id = int.Parse(Reader[0].ToString());
        //        if (PreId + 1 != Id)
        //        {
        //            Id = PreId;
        //            break;
        //        }
        //        PreId = Id;
        //    }
        //    Id++;
        //    Reader.Close();
        //    CloseSQLConnection();
        //    return Id;
        //}

        //public long GetLastId(string TableName, string colName = "Id")
        //{
        //    long Id = 0;
        //    string query = "SELECT " + colName + " FROM " + TableName;
        //    Command = new SqlCommand(query, GetSQLConnection());
        //    OpenSQLConnection();
        //    SqlDataReader Reader = Command.ExecuteReader();
        //    while (Reader.Read())
        //    {
        //        Id = int.Parse(Reader[0].ToString());
        //    }
        //    Reader.Close();
        //    CloseSQLConnection();
        //    return Id;
        //}
        //public void BindDropDownList(DropDownList ddl, string query, string text, string value, string defaultText)
        //{
        //    SqlCommand cmd = new SqlCommand(query);

        //    cmd.Connection = GetSQLConnection();
        //    //con.Open();
        //    ddl.DataSource = cmd.ExecuteReader();
        //    ddl.DataTextField = text;
        //    ddl.DataValueField = value;
        //    ddl.DataBind();
        //    //con.Close();

        //    ddl.Items.Insert(0, new ListItem(defaultText, "0"));
        //}

        public void Execute_Query(string query)
        {
            OpenSQLConnection();
            Command = new SqlCommand(query, GetSQLConnection());
            Command.ExecuteNonQuery();
            CloseSQLConnection();
        }

        //public void ExecuteParameterizedQuery(string query, params object[] parameter)
        //{
        //    string[] temp = query.Split('@');
        //    int Count = temp.Length - 1;
        //    if (parameter.Length != Count)
        //    {
        //        return;
        //    }
        //    string[] Parameter = new string[Count];

        //    for (int i = 1; i < temp.Length; i++)
        //    {
        //        string[] stemp = temp[i].Split(' ');
        //        string[] ctemp = stemp[0].Split(',');
        //        string[] cbtemp = ctemp[0].Split(')');
        //        string[] obtemp = cbtemp[0].Split('(');
        //        string[] sctemp = obtemp[0].Split(';');
        //        Parameter[i - 1] = "@" + sctemp[0];
        //    }

        //    OpenSQLConnection();
        //    Command = new SqlCommand();
        //    Command.Connection = GetSQLConnection();
        //    Command.CommandType = CommandType.Text;
        //    Command.CommandText = query;
        //    for (int i = 0;i<Count;i++)
        //    {
        //        Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
        //    }
        //    Command.ExecuteNonQuery();
        //    CloseSQLConnection();
        //}  

        public void ExecuteProcedure(string procedure)
        {
            OpenSQLConnection();
            Command = new SqlCommand();
            Command.Connection = GetSQLConnection();
            Command.CommandType = System.Data.CommandType.StoredProcedure;
            Command.CommandText = procedure;
            Command.ExecuteNonQuery();
            CloseSQLConnection();
        }

        public void ExecuteParameterizedProcedure(string procedure, string stringParameters, params object[] parameter)
        {
            string[] temp = stringParameters.Split('@');
            int Count = temp.Length - 1;
            if (parameter.Length != Count)
            {
                return;
            }
            string[] Parameter = new string[Count];

            for (int i = 1; i < temp.Length; i++)
            {
                string[] stemp = temp[i].Split(' ');
                string[] ctemp = stemp[0].Split(',');
                string[] cbtemp = ctemp[0].Split(')');
                string[] obtemp = cbtemp[0].Split('(');
                string[] sctemp = obtemp[0].Split(';');
                Parameter[i - 1] = "@" + sctemp[0];
            }
            OpenSQLConnection();
            Command = new SqlCommand();
            Command.Connection = GetSQLConnection();
            Command.CommandType = System.Data.CommandType.StoredProcedure;
            Command.CommandText = procedure;
            for (int i = 0; i < Count; i++)
            {
                //if (string.IsNullOrEmpty(parameter[i].ToString()))
                //    if (parameter[i].Equals(null))
                //    {
                //        Command.Parameters.AddWithValue(Parameter[i], System.DBNull.Value);
                //    }
                //    else
                //    {
                //        Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
                //    }

                //Command.Parameters.AddWithValue(Parameter[i], parameter[i]);

                if (parameter[i] != null)
                {
                    Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
                }
                else
                {
                    Command.Parameters.AddWithValue(Parameter[i], System.DBNull.Value);
                }

            }
            Command.ExecuteNonQuery();

            CloseSQLConnection();
        }

        public string QIUD(SqlCommand cmd)
        {
            try
            {
                OpenSQLConnection();
                cmd.Connection = GetSQLConnection();
                //OpenSQLConnection();
                cmd.ExecuteNonQuery();
                CloseSQLConnection();
                return "Done";
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public DataTable Return_DataTable_Query(string query)
        {
            OpenSQLConnection();
            Command = new SqlCommand();
            Command.Connection = GetSQLConnection();
            Command.CommandType = System.Data.CommandType.Text;
            Command.CommandText = query;
            Adapter = new SqlDataAdapter(Command);
            Adapter.SelectCommand.CommandTimeout = 120;  // seconds
            DataSet Set = new DataSet();
            Adapter.Fill(Set);
            CloseSQLConnection();
            return Set.Tables[0];
        }
        public SqlDataReader Return_DataReader_Query(string query)
        {
            OpenSQLConnection();
            Command = new SqlCommand();
            Command.Connection = GetSQLConnection();
            Command.CommandType = System.Data.CommandType.Text;
            Command.CommandText = query;
            SqlDataReader reader;
            reader = Command.ExecuteReader();
            CloseSQLConnection();
            return reader;
        }
        //public DataSet Return_DataSet_Query(string Query)
        //{
        //    try
        //    {
        //        OpenSQLConnection();
        //        ds = new DataSet();
        //        ODBC_Cmd = new OdbcCommand(Query, GetCon());
        //        ODBC_da = new OdbcDataAdapter(ODBC_Cmd);
        //        ODBC_da.Fill(ds.Tables[0]);
        //        return ds;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        //public DataTable ReturnParameterizedDataTableQuery(string query, params object[] parameter)
        //{
        //    string[] temp = query.Split('@');
        //    int Count = temp.Length - 1;
        //    if (parameter.Length != Count)
        //    {
        //        return null;
        //    }
        //    string[] Parameter = new string[Count];

        //    for (int i = 1; i < temp.Length; i++)
        //    {
        //        string[] stemp = temp[i].Split(' ');
        //        string[] ctemp = stemp[0].Split(',');
        //        string[] cbtemp = ctemp[0].Split(')');
        //        string[] obtemp = cbtemp[0].Split('(');
        //        string[] sctemp = obtemp[0].Split(';');
        //        Parameter[i - 1] = "@" + sctemp[0];
        //    }
        //    OpenSQLConnection();
        //    Command = new SqlCommand();
        //    Command.Connection = GetSQLConnection();
        //    Command.CommandType = System.Data.CommandType.Text;
        //    Command.CommandText = query;
        //    for (int i = 0; i < Count; i++)
        //    {
        //        Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
        //    }
        //    Adapter = new SqlDataAdapter(Command);
        //    DataSet Set = new DataSet();
        //    Adapter.Fill(Set);
        //    CloseSQLConnection();
        //    return Set.Tables[0];
        //}

        public DataTable ReturnDataTableProcedure(string procedure)
        {
            OpenSQLConnection();
            Command = new SqlCommand();
            Command.Connection = GetSQLConnection();
            Command.CommandType = System.Data.CommandType.StoredProcedure;
            Command.CommandText = procedure;
            Adapter = new SqlDataAdapter(Command);
            Adapter.SelectCommand.CommandTimeout = 120;  // seconds
            DataSet Set = new DataSet();
            Adapter.Fill(Set);
            CloseSQLConnection();
            return Set.Tables[0];
        }

        public DataTable ReturnParameterizedDataTableProcedure(string procedure, string stringParameters, params object[] parameter)
        {
            string[] temp = stringParameters.Split('@');
            int Count = temp.Length - 1;
            if (parameter.Length != Count)
            {
                return null;
            }
            string[] Parameter = new string[Count];

            for (int i = 1; i < temp.Length; i++)
            {
                string[] stemp = temp[i].Split(' ');
                string[] ctemp = stemp[0].Split(',');
                string[] cbtemp = ctemp[0].Split(')');
                string[] obtemp = cbtemp[0].Split('(');
                string[] sctemp = obtemp[0].Split(';');
                Parameter[i - 1] = "@" + sctemp[0];
            }
            OpenSQLConnection();
            Command = new SqlCommand();
            Command.Connection = GetSQLConnection();
            Command.CommandType = System.Data.CommandType.StoredProcedure;
            Command.CommandText = procedure;
            for (int i = 0; i < Count; i++)
            {
                Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
            }
            Adapter = new SqlDataAdapter(Command);
            Adapter.SelectCommand.CommandTimeout = 120;  // seconds
            DataSet Set = new DataSet();
            Adapter.Fill(Set);
            CloseSQLConnection();
            return Set.Tables[0];
        }

        //public bool CheckRecord(string query)
        //{
        //    bool check = false;
        //    OpenSQLConnection();
        //    Command = new SqlCommand(query, GetSQLConnection());
        //    SqlDataReader Reader = Command.ExecuteReader();
        //    if (Reader.Read())
        //    {
        //        check = true;
        //    }
        //    CloseSQLConnection();
        //    return check;
        //}

        //public bool CheckParameterizedRecord(string query, params object[] parameter)
        //{
        //    string[] temp = query.Split('@');
        //    int Count = temp.Length - 1;
        //    if (parameter.Length != Count)
        //    {
        //        return false;
        //    }
        //    string[] Parameter = new string[Count];

        //    for (int i = 1; i < temp.Length; i++)
        //    {
        //        string[] stemp = temp[i].Split(' ');
        //        string[] ctemp = stemp[0].Split(',');
        //        string[] cbtemp = ctemp[0].Split(')');
        //        string[] obtemp = cbtemp[0].Split('(');
        //        string[] sctemp = obtemp[0].Split(';');
        //        Parameter[i - 1] = "@" + sctemp[0];
        //    }
        //    bool check = false;
        //    OpenSQLConnection();
        //    Command = new SqlCommand(query, GetSQLConnection());
        //    for (int i = 0; i < Count; i++)
        //    {
        //        Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
        //    }
        //    SqlDataReader Reader = Command.ExecuteReader();
        //    if (Reader.Read())
        //    {
        //        check = true;
        //    }
        //    CloseSQLConnection();
        //    return check;
        //}

        //public object FetchResult(string query, int Index = 0)
        //{
        //    object result = new object();
        //    OpenSQLConnection();
        //    Command = new SqlCommand(query, GetSQLConnection());

        //    SqlDataReader Reader = Command.ExecuteReader();
        //    if (Reader.Read())
        //    {
        //        result= Reader[Index];
        //    }
        //    CloseSQLConnection();
        //    return result;
        //}

        //public object FetchParameterizedResult(string query, int Index = 0, params object[] parameter)
        //{
        //    string[] temp = query.Split('@');
        //    int Count = temp.Length - 1;
        //    if (parameter.Length != Count)
        //    {
        //        return null;
        //    }
        //    string[] Parameter = new string[Count];

        //    for (int i = 1; i < temp.Length; i++)
        //    {
        //        string[] stemp = temp[i].Split(' ');
        //        string[] ctemp = stemp[0].Split(',');
        //        string[] cbtemp = ctemp[0].Split(')');
        //        string[] obtemp = cbtemp[0].Split('(');
        //        string[] sctemp = obtemp[0].Split(';');
        //        Parameter[i - 1] = "@" + sctemp[0];
        //    }
        //    object result = new object();
        //    OpenSQLConnection();
        //    Command = new SqlCommand(query, GetSQLConnection());
        //    for (int i = 0; i < Count; i++)
        //    {
        //        Command.Parameters.AddWithValue(Parameter[i], parameter[i]);
        //    }
        //    SqlDataReader Reader = Command.ExecuteReader();
        //    if (Reader.Read())
        //    {
        //        result = Reader[Index];
        //    }
        //    CloseSQLConnection();
        //    return result;
        //}
    }
}
