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
using QuanLyMauKiemNghiem.Object;
using QuanLyMauKiemNghiem.Model;
namespace QuanLyMauKiemNghiem.View
{
    public partial class frmThemKhachHang : DevExpress.XtraEditors.XtraForm
    {
        Khach_Hang_Obj khObj = new Khach_Hang_Obj();
        Khach_Hang_Mod khMod = new Khach_Hang_Mod();
        LoadData ld = new LoadData();
        

        public frmThemKhachHang()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            khObj.Makh = ld.AutoID("KH-", "SELECT MAX(RIGHT(MAKH,7)) from KHACH_HANG");
            khObj.Tenkh = textEdit1.Text;
            khObj.Diachi = textEdit2.Text;
            khObj.Masothue = textEdit3.Text;
            khObj.Sdt = textEdit4.Text;
            khObj.Email = textEdit5.Text;
            khObj.Sofax = textEdit6.Text;
            khObj.Nguoi_lh = textEdit7.Text;
            khObj.Sdt_lh = textEdit8.Text;
            khObj.Email_lh = textEdit9.Text;
            if (khObj.Tenkh == "" || khObj.Diachi == "" || khObj.Sdt == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ thông tin");
                return;
            }
            else
            {
                if (khMod.addData(khObj))
                {
                    XtraMessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    frmThemHopDong.makh = khObj.Makh;
                    this.Close();
                }
                else
                {
                    XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }


    }

}