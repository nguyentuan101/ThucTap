using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
    class ConnectToSQL
    {
        public static string strServer, strUser, strPass;
        private SqlConnection Conn;
        private SqlCommand _cmd;
        public SqlCommand CMD
        {
            get { return _cmd; }
            set { _cmd = value; }
        }

        public SqlConnection Connection { get { return Conn; } }
        private string error;

        public string Error
        {
            get { return error; }
            set { error = value; }
        }       
    
        public ConnectToSQL()
        {
            strServer = Properties.Settings.Default.Server;
            strUser = Properties.Settings.Default.User;
            strPass = Properties.Settings.Default.Pass;
            string strCon = "Server=" + strServer + ";Database=QLMKN2;Integrated Security = false; UID="+strUser+"; PWD="+strPass;
            Conn = new SqlConnection(strCon);
            Conn.Open();

        }



        public bool OpenConn()
        {
            try
            {
                if (Conn.State == ConnectionState.Closed)
                    Conn.Open();
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            return true;
        }

        public bool CloseConn()
        {
            try
            {
                if (Conn.State == ConnectionState.Open)
                    Conn.Close();
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            return true;
        }


    }
}
