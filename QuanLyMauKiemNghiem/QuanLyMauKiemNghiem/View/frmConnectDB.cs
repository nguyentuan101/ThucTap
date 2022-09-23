using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using QuanLyMauKiemNghiem.Model;
using System.Data.Sql;


namespace QuanLyMauKiemNghiem.View
{
    public partial class frmConnectDB : DevExpress.XtraEditors.XtraForm
    {

        //private Login fConect;

        //public frmConnectDB(Login fm)
        //    : this()
        //{
        //    fConect = fm;
        //}


        public frmConnectDB()
        {
            InitializeComponent();
        }


        private void btnConnect_Click(object sender, EventArgs e)
        {
            string server = Properties.Settings.Default.Server;
            string user = Properties.Settings.Default.User;
            string pass = Properties.Settings.Default.Pass;

            Properties.Settings.Default.Server = cboServer.SelectedValue.ToString();           
            Properties.Settings.Default.User = txtUser.Text.Trim();
            Properties.Settings.Default.Pass = txtPass.Text.Trim();
            Properties.Settings.Default.Save();            
            try
            {
                ConnectToSQL con = new ConnectToSQL();
                SqlCommand cmd = new SqlCommand();
                MessageBox.Show("Kết nối thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //frmMain.connect = 1;
                //this.Close();
                this.Hide();
                //frmMain FM = new frmMain();
                //if (FM != null)
                //    FM.Close();
                Login fL = new Login();
                fL.ShowDialog();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Kết nối thất bại: "+ex, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Properties.Settings.Default.Server = server;
                Properties.Settings.Default.User = user;
                Properties.Settings.Default.Pass = pass;
                Properties.Settings.Default.Save(); 
            }                     
        }


        private void frmConnectDB_Load(object sender, EventArgs e)
        {           
            
            loadCboServer();
        }    


        //Load cac ServerName len combobox
        private static DataTable DisplayData(DataTable table)
        {
            DataTable tb = new DataTable();
            tb.Columns.Add("ID", typeof(string));
            tb.Columns.Add("ServerName", typeof(string));
            foreach (DataRow row in table.Rows)
            {
                tb.Rows.Add(row["ServerName"] + "\\" + row["InstanceName"], row["ServerName"]);
            }
            return tb;
        }


       void loadCboServer()
        {
            SqlDataSourceEnumerator instance = SqlDataSourceEnumerator.Instance;
            DataTable table = instance.GetDataSources();
            cboServer.DataSource = DisplayData(table);
            cboServer.DisplayMember = "ID";
            cboServer.ValueMember = "ID";   
        }


       private void btnCancel_Click(object sender, EventArgs e)
       {
           this.Close();
       }


   


    }
}