using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
    public class LoadData
    {
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();

        //Tăng mã
        public string AutoID(string kyhieu,string select)
        {
            DataTable dt = new DataTable();
            cmd.CommandText = select;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = con.Connection;
            con.OpenConn();
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt);
            string ma = "";
            DataRow ROW = dt.Rows[0];
            
            if (ROW[0].ToString() == "")
                ma = ma + "0000000";
            else
            {
                int k;
                ma = "";
                k = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][0].ToString().Substring(0, 7));
                k = k + 1;
                if (k < 10)
                    ma = ma + "000000";
                else
                {
                    if (k < 100)
                        ma = ma + "00000";
                    else
                    {
                        if (k < 1000)
                            ma = ma + "0000";
                        else
                        {
                            if (k < 10000)
                                ma = ma + "000";
                            else
                            {
                                if (k < 100000)
                                    ma = ma + "00";
                                else
                                    ma = ma + "0";
                            }
                        }
                    }
                }
                ma = ma + k.ToString();
            }
            return kyhieu + ma;
        }

        public DataTable GetData(string sql)
        {
            DataTable dt = new DataTable();
            cmd.CommandText = sql;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = con.Connection;
            try
            {
                con.OpenConn();
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                con.CloseConn();
            }
            catch (Exception ex)
            {
                string mex = ex.Message;
                cmd.Dispose();
                con.CloseConn();
            }
            return dt;
        }

        //Gọi ham them,sua,xoa dữ liệu
        public bool AddData(string sql)
        {
             try
            {
                cmd.CommandText = sql;
                cmd.CommandType = CommandType.Text;
                cmd.Connection = con.Connection;
                con.OpenConn();
                cmd.ExecuteNonQuery();
                con.CloseConn();
                return true;
            }
            catch (Exception ex)
            {
                string mex = ex.Message;
                cmd.Dispose();
                con.CloseConn();

            }
            return false;
        }


    }
}
