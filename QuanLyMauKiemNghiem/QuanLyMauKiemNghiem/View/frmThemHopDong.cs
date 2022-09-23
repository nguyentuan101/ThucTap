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
using QuanLyMauKiemNghiem.Object;

namespace QuanLyMauKiemNghiem.View
{
    public partial class frmThemHopDong : DevExpress.XtraEditors.XtraForm
    {
        HopDong_Mod hdMod = new HopDong_Mod();
        Hop_Dong_Obj hdObj = new Hop_Dong_Obj();
        LoadData l = new LoadData();
        public static string makh;

        public frmThemHopDong()
        {
            InitializeComponent();            
        }



        void loadkhachhang()
        {
            searchLookUpEdit1.Properties.DataSource = hdMod.GetKhachHang();
            searchLookUpEdit1.Properties.ValueMember = "MAKH";
            searchLookUpEdit1.Properties.DisplayMember = "TENKH";


        }

        private void frmThemHopDong_Load(object sender, EventArgs e)
        {
            labelControl1.Text = "";
            loadkhachhang();

        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (searchLookUpEdit1.Text == "")
            {
                labelControl1.Text = "Chưa chọn khách hàng!";
                return;
            }
            if (textEdit1.Text == "" || textEdit2.Text == "")
            {
                labelControl1.Text = "Vui lòng điền mã hợp đồng và tên hợp đồng!";
                return;
            }
            else
            {
                hdObj.Mahd = textEdit1.Text;
                if (hdMod.KTHopDong(hdObj) > 0)
                {
                    labelControl1.Text = "Mã hợp đồng đã tồn tại, Vui lòng kiểm tra lại!";
                    return;
                }
                hdObj.Tenhd = textEdit2.Text;
                hdObj.Makh = searchLookUpEdit1.EditValue.ToString();
                hdObj.Tenkh = searchLookUpEdit1.Text;
                hdObj.Thoihan = textEdit4.Text;
                hdObj.Noidung = memoEdit1.Text;
                hdObj.Nguoichiutn = textEdit3.Text;
                hdObj.Tinhtrang = textEdit5.Text;
                if (hdMod.ThemHopDong(hdObj))
                {
                    XtraMessageBox.Show("Thêm mới hợp đồng thành công!","Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    frmMain.tempMahd = hdObj.Mahd;
                    this.Close();
                }
                else
                {
                    labelControl1.Text = "";
                    XtraMessageBox.Show("Thêm mới hợp đồng thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            frmThemKhachHang fm = new frmThemKhachHang();
            fm.ShowDialog();
            loadkhachhang();
            searchLookUpEdit1.EditValue = makh;
        }
    }
}