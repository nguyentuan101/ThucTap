using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using QuanLyMauKiemNghiem.Model;
using System.Data.SqlClient;
using System.IO;
namespace QuanLyMauKiemNghiem.View
{
    public partial class Login : DevExpress.XtraEditors.XtraForm
    {
     
        public static string tenUser;
        public static string Mkhau;
       
        public Login()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            
            this.Close();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            //txtTenUser.Text = "admin";
            //txtMatKhau.Text = "123";
        }
        public static int login = 0;
        private void btnConnect_Click(object sender, EventArgs e)
        {
                       
            try
            {
                ConnectToSQL con = new ConnectToSQL();
                SqlCommand cmd = new SqlCommand();
                DataTable dt = new DataTable();
                cmd.CommandText = "select * from [QLMKN2].[dbo].[NGUOI_DUNG] where TENUSER ='" + txtTenUser.Text.Trim() + "' and MATKHAU = '" + txtMatKhau.Text.Trim() + "'";
                cmd.CommandType = CommandType.Text;
                cmd.Connection = con.Connection;
                con.OpenConn();
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (dt.Rows.Count >0)
                {
                    MessageBox.Show("Đăng nhập thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tenUser = txtTenUser.Text.Trim();
                    Mkhau = txtMatKhau.Text.Trim();
                    this.Hide();
                   frmMain kh = new frmMain();                    
                   kh.ShowDialog();
                   this.Close();
                }
                    
                else
                    MessageBox.Show("Tên đăng nhập hoặc mật khẩu sai! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Lỗi kết nối đến cơ sở dữ liệu! ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Hide();
                frmConnectDB fm = new frmConnectDB();
                fm.ShowDialog();

            }
           
        }

        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}