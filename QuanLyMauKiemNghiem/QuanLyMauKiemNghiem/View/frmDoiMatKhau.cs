using QuanLyMauKiemNghiem.Object;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QuanLyMauKiemNghiem.Model;

namespace QuanLyMauKiemNghiem.View
{
    public partial class frmDoiMatKhau : Form
    {
        Nguoi_Dung_Obj ndObj = new Nguoi_Dung_Obj();
        Nguoi_dung_Mod ndMod = new Nguoi_dung_Mod();

        public frmDoiMatKhau()
        {
            InitializeComponent();
        }

        
        public void loadMK()
        {

            string mkhaucu = Login.Mkhau;
            string tenDangNhap = txtTenTK.Text;

            string matKhau = txtMKHienTai.Text;
            string matKhauMoi = txtMKMoi.Text;
            string nhapLaiMK = txtNhapLaiMK.Text;
            if (!matKhauMoi.Equals(nhapLaiMK))
            {
                MessageBox.Show("Mật khẩu không khớp!");
            }

            else if (!matKhau.Equals(mkhaucu))
            {
                MessageBox.Show("Mật Khẩu Không Đúng!");
            }

            else
            {
                ndObj.Matkhau = matKhauMoi;
                ndObj.Tenuser = tenDangNhap;
                ndObj.Matkhaucu = matKhau;

                if (ndMod.DoiMK(ndObj))
                {
                    MessageBox.Show("Cập Nhật Thông Tin Thành Công");
                    Application.Restart();
                }

                else
                {
                    MessageBox.Show("Mật Khẩu Không Đúng!");
                }
            }

        }

        private void frmDoiMatKhau_Load(object sender, EventArgs e)
        {
            txtTenTK.Text = Login.tenUser;


        }

        private void btnXacnhan_Click(object sender, EventArgs e)
        {
            loadMK();
        }
    }
}

