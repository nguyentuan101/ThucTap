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
using QuanLyMauKiemNghiem.View;

namespace QuanLyMauKiemNghiem.View
{
    public partial class frmThemChiTieu : DevExpress.XtraEditors.XtraForm
    {
        //
        //Thêm mới chỉ định bị lỗi
        //
        CapNhatMau_Mod cnmMod = new CapNhatMau_Mod();
        Nhan_Mau_Mod nhanmauMod = new Nhan_Mau_Mod();
        Phieu_Nhan_Mau_Obj nhanmauObj = new Phieu_Nhan_Mau_Obj();
        Gia_Chi_Tieu_Obj gctObj = new Gia_Chi_Tieu_Obj();
        Phuong_Phap_Obj ppObj = new Phuong_Phap_Obj();
        ChiTieu_Mod ctMod = new ChiTieu_Mod();
        NhapKetQua_Mod nkqMod = new NhapKetQua_Mod();
        //string mahd_cnm = frmMain.mahd_cnm;
        // string tenmau_cnm=frmMain.tenmau_cnm;
        // string ngay_cnm = frmMain.ngay_cnm;            
        string macd_18;

        public static string tempMAMAU_CNM;

        public frmThemChiTieu()
        {
            InitializeComponent();
        }


        void disEnable(bool e)
        {
            textEdit3.ReadOnly = e;
            cboNoiKiem_18.ReadOnly = e;
            cboLoaiGia.ReadOnly = e;
            txtDonGia.ReadOnly = e;
            dgvDSCT_18.Enabled = e;
            gc_ChiTieu.Enabled = e;
            gc_ChiDinh.Enabled = e;
            if (e)
            {
                layoutControlItem24.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem25.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else
            {
                layoutControlItem24.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem25.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                emptySpaceItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            }
                
        }
        void loadPP()
        {
            textEdit3.Properties.DataSource = cnmMod.GetPhuongPhap();
            textEdit3.Properties.ValueMember = "MAPP";
            textEdit3.Properties.DisplayMember = "TENPP";
            cboPPChiTieu.DataSource = cnmMod.GetPhuongPhap();
            cboPPChiTieu.ValueMember = "MAPP";
            cboPPChiTieu.DisplayMember = "TENPP";
        }
        void load_CNCTK()
        {
            gc_ChiDinh.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gc_ChiTieu.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            dgvDSCT_18.DataSource = cnmMod.GetDSChiTieu(tempMAMAU_CNM);
            dgvChiTieu_18.DataSource = cnmMod.GetChiTieu(tempMAMAU_CNM);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = cnmMod.GetData(tempMAMAU_CNM);
            
            txtMaHopDong.Text = dt.Rows[0]["MAHD"].ToString();
            txtKH.Text = dt.Rows[0]["TENKH"].ToString();
            txtJobNo.Text = dt.Rows[0]["JOBNO"].ToString();
            txtMaMau.Text = dt.Rows[0]["MAMAU"].ToString();
            txtTenMau_18.Text = dt.Rows[0]["TENMAU"].ToString();

            cboNoiKiem_18.Properties.DataSource = cnmMod.GetNoiKiem();
            cboNoiKiem_18.Properties.ValueMember = "MANK";
            cboNoiKiem_18.Properties.DisplayMember = "TENNK";
            cboNoiKiem_18.EditValue = "";
            loadPP();           


            cboNoiKiem_18.DataBindings.Clear();
            cboNoiKiem_18.DataBindings.Add("EditValue", dgvDSCT_18.DataSource, "MANK");
            txtDonGia.DataBindings.Clear();
            txtDonGia.DataBindings.Add("Text", dgvDSCT_18.DataSource, "DONGIA");
            layoutControlItem21.DataBindings.Clear();
            layoutControlItem21.DataBindings.Add("Text", dgvDSCT_18.DataSource, "MACT");
            textEdit2.DataBindings.Clear();
            textEdit2.DataBindings.Add("Text", dgvDSCT_18.DataSource, "TENCT");
            textEdit3.DataBindings.Clear();
            textEdit3.DataBindings.Add("EditValue", dgvDSCT_18.DataSource, "MAPP");
            layoutControlItem28.DataBindings.Clear();
            layoutControlItem28.DataBindings.Add("Text", dgvDSCT_18.DataSource, "MAPHIEU_YCPT");
            textEdit6.DataBindings.Clear();
            textEdit6.DataBindings.Add("Text", dgvDSCT_18.DataSource, "KETQUA_PT");
            layoutControlItem29.DataBindings.Clear();
            layoutControlItem29.DataBindings.Add("Text", dgvDSCT_18.DataSource, "TRANGTHAI_CHUYEN");
            

            cboLoaiGia.Properties.DataSource = cnmMod.GetLoaiGia(layoutControlItem21.Text);
            cboLoaiGia.Properties.ValueMember = "GIA";
            cboLoaiGia.Properties.DisplayMember = "TENPHANLOAIGIA";
            disEnable(true);
           
        }
        private void simpleButton1_Click(object sender, EventArgs e)//Them chỉ dinh
        {
            if (cboChiDinh_18.Text == "")
            {
                MessageBox.Show("Vui lòng chọn chỉ định");
                return;
            }

            DataRow row = nhanmauMod.CopyMau(tempMAMAU_CNM).Rows[0];
            nhanmauObj.Time_nhanmaudk = Convert.ToDateTime(row["TIME_NHANMAU_DK"]).ToString("yyyy/MM/dd");
            nhanmauObj.Mahd = row["MAHD"].ToString();
            nhanmauObj.Tenmau = row["TENMAU"].ToString();
            nhanmauObj.Motamau = row["MOTAMAU"].ToString();
            nhanmauObj.Khoiluongmau = row["KHOILUONGMAU"].ToString();
            nhanmauObj.Tinhtrangmau = row["TINHTRANGMAU"].ToString();
            nhanmauObj.Ttkhcungcap = row["TTKHCUNGCAP"].ToString();
            nhanmauObj.Nguonmau = row["NGUONMAU"].ToString();
            nhanmauObj.Seal = row["SEAL"].ToString();
            nhanmauObj.Nguoinhanmau = row["NGUOINHANMAU"].ToString();
            nhanmauObj.TempMaMau = row["temp_MAMAU"].ToString();
            nhanmauObj.Mact = row["MACT"].ToString();
            nhanmauObj.Maphieu = row["MAPHIEU_YCPT"].ToString();


            if (row["TIME_TRABCPT_DK"].ToString() != "")
                nhanmauObj.Time_trabcpt_dk = Convert.ToDateTime(row["TIME_TRABCPT_DK"]).ToString("yyyy/MM/dd");
            else
                nhanmauObj.Time_trabcpt_dk = "";


            string text = nhanmauMod.Check_UOA(nhanmauObj).Rows[0]["MACT"].ToString();

            System.Data.DataTable dt = nhanmauMod.GetChiTieu_ChiDinh(macd_18, nhanmauObj);

            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("Các chỉ tiêu đã có trong mẫu " + nhanmauObj.Tenmau);
                return;
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {


                nhanmauObj.Mact = dt.Rows[i]["MACT"].ToString();
                nhanmauObj.Tenct = dt.Rows[i]["TENCT"].ToString();
                nhanmauObj.Manenmau = dt.Rows[i]["MANENMAU"].ToString();
                nhanmauObj.Tennenmau = dt.Rows[i]["TENNENMAU"].ToString();
                nhanmauObj.Lod = dt.Rows[i]["LOD"].ToString();
                nhanmauObj.Madv = dt.Rows[i]["MADV"].ToString();
                nhanmauObj.Macd = dt.Rows[i]["MACD"].ToString();
                nhanmauObj.Tencd = dt.Rows[i]["TENCD"].ToString();
                if (nhanmauObj.Madv == "")
                    nhanmauObj.Madv = "NULL";
                nhanmauObj.Tendv = dt.Rows[i]["TENDV"].ToString();

                nhanmauObj.Dongia = dt.Rows[i]["GIA"].ToString();
                if (nhanmauObj.Dongia == "")
                    nhanmauObj.Dongia = "0";
                nhanmauObj.Mapp = dt.Rows[i]["MAPP"].ToString();
                if (nhanmauObj.Mapp == "")
                    nhanmauObj.Mapp = "NULL";
                nhanmauObj.Tenpp = dt.Rows[i]["TENPP"].ToString();
                nhanmauObj.Nhomct = dt.Rows[i]["NHOMCHITIEU"].ToString();
                nhanmauObj.Maqc = dt.Rows[i]["MAQC"].ToString();
                if (nhanmauObj.Maqc == "")
                    nhanmauObj.Maqc = "NULL";
                nhanmauObj.Tenqc = dt.Rows[i]["TENQC"].ToString();

                if (text.Equals(""))
                {
                    nhanmauMod.updateChiTieuDau(nhanmauObj);
                    text = "1";
                }
                else
                {
                    if (nhanmauMod.KiemTraChiTieu(nhanmauObj).Rows.Count == 0)
                        nhanmauMod.addDataNew(nhanmauObj);
                }
            }

            XtraMessageBox.Show("Đã thêm chỉ định cho mẫu " + nhanmauObj.Tenmau);
            macd_18 = "";
            gc_ChiDinh.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gc_ChiTieu.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
        }

        private void btnThemMoiChiDinh_Click(object sender, EventArgs e)//thêm mới chỉ dinh
        {
            gc_ChiDinh.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            gc_ChiTieu.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
           

            cboChiDinh_18.Properties.DataSource = cnmMod.GetChiDinh();
            cboChiDinh_18.Properties.ValueMember = "MACD";
            cboChiDinh_18.Properties.DisplayMember = "TENCD";
        }

        private void frmThemChiTieu_Load(object sender, EventArgs e)
        {
            load_CNCTK();
        }

        private void simpleButton2_Click(object sender, EventArgs e)//
        {
            gc_ChiDinh.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gc_ChiTieu.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
        }

        private void btnCapNhat_Click(object sender, EventArgs e)//
        {
            int[] chitieu = gridView45.GetSelectedRows();
            if (chitieu.Length == 0)
            {
                XtraMessageBox.Show("Vui lòng chọn chỉ tiêu!");
                return;
            }

            //if (layoutControlItem29.Text == "Đã xuất" || layoutControlItem29.Text == "Hoàn thành" || layoutControlItem29.Text == "Đã chuyển" || layoutControlItem29.Text == "Kiểm lại")
            //{
            //    XtraMessageBox.Show("Mẫu đang kiểm hoặc đã xuất kết quả,không thể thêm chỉ tiêu");
            //    return;
            //}

            int[] k = gridView44.GetSelectedRows();
            
            DataRow row = gridView44.GetDataRow(k[0]);
            nhanmauObj.Time_nhanmaudk = Convert.ToDateTime(row["TIME_NHANMAU_DK"]).ToString("yyyy/MM/dd");
            nhanmauObj.Mahd = row["MAHD"].ToString();
            nhanmauObj.Tenmau = row["TENMAU"].ToString();
            nhanmauObj.Motamau = row["MOTAMAU"].ToString();
            nhanmauObj.Khoiluongmau = row["KHOILUONGMAU"].ToString();
            nhanmauObj.Tinhtrangmau = row["TINHTRANGMAU"].ToString();
            nhanmauObj.Ttkhcungcap = row["TTKHCUNGCAP"].ToString();
            nhanmauObj.Nguonmau = row["NGUONMAU"].ToString();
            nhanmauObj.Seal = row["SEAL"].ToString();
            nhanmauObj.Nguoinhanmau = row["NGUOINHANMAU"].ToString();
            nhanmauObj.TempMaMau = row["temp_MAMAU"].ToString();
            nhanmauObj.Mact = row["MACT"].ToString();
            nhanmauObj.Mamau = row["MAMAU"].ToString();          
            
            nhanmauObj.Jobno = row["JOBNO"].ToString();
            nhanmauObj.Maphieu = row["MAPHIEU_YCPT"].ToString();

            if (row["TIME_NHANMAU_TT"].ToString() != "")
                nhanmauObj.Time_nhanmau_tt = Convert.ToDateTime(row["TIME_NHANMAU_TT"]).ToString("yyyy/MM/dd");
            else
                nhanmauObj.Time_nhanmau_tt = "";

            if (row["TIME_TRABCPT_DK"].ToString() != "")
                nhanmauObj.Time_trabcpt_dk = Convert.ToDateTime(row["TIME_TRABCPT_DK"]).ToString("yyyy/MM/dd");
            else
                nhanmauObj.Time_trabcpt_dk = "";

            string count = row["MACT"].ToString();

            for (int i = 0; i < chitieu.Length; i++)
            {
                DataRow row1 = gridView45.GetDataRow(chitieu[i]);
                nhanmauObj.Mact = row1["MACT"].ToString();
                nhanmauObj.Tenct = row1["TENCT"].ToString();
                nhanmauObj.Manenmau = row1["MANENMAU"].ToString();
                nhanmauObj.Tennenmau = row1["TENNENMAU"].ToString();
                nhanmauObj.Lod = row1["LOD"].ToString();
                nhanmauObj.Madv = row1["MADV"].ToString();
                if (nhanmauObj.Madv == "")
                    nhanmauObj.Madv = "NULL";
                nhanmauObj.Tendv = row1["TENDV"].ToString();

                nhanmauObj.Dongia = row1["GIA"].ToString();
                if (nhanmauObj.Dongia == "")
                    nhanmauObj.Dongia = "0";
                nhanmauObj.Mapp = row1["MAPP"].ToString();

                if (nhanmauObj.Mapp == "")
                    nhanmauObj.Mapp = "NULL";
                if (nkqMod.getTenPhuongPhap(nhanmauObj.Mapp).Rows.Count > 0)
                {
                    DataRow row4 = nkqMod.getTenPhuongPhap(nhanmauObj.Mapp).Rows[0];
                    nhanmauObj.Tenpp = row4[0].ToString();
                }
                else
                    nhanmauObj.Tenpp = "";
                
                nhanmauObj.Nhomct = row1["NHOMCHITIEU"].ToString();
                nhanmauObj.Maqc = row1["MAQC"].ToString();
                if (nhanmauObj.Maqc == "")
                    nhanmauObj.Maqc = "NULL";
                nhanmauObj.Tenqc = row1["TENQC"].ToString(); 
                
                if (count == "")
                {                    
                    cnmMod.updateChiTieuDau(nhanmauObj);
                    count = "1";
                }
                else
                {
                    if (nhanmauMod.KiemTraChiTieu(nhanmauObj).Rows.Count == 0)
                        cnmMod.addDataNew(nhanmauObj);
                           
                }

            }
            XtraMessageBox.Show("Đã thêm chỉ tiêu kiểm cho mẫu " + nhanmauObj.Tenmau);
            load_CNCTK();
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            load_CNCTK();
        }

      

        private void txtDonGia_TextChanged(object sender, EventArgs e)
        {
            cboLoaiGia.EditValue = txtDonGia.Text;
        }

        private void cboLoaiGia_EditValueChanged(object sender, EventArgs e)
        {
            if (cboLoaiGia.EditValue == null)
                txtDonGia.Text = "0";
            else
                txtDonGia.Text = cboLoaiGia.EditValue.ToString();
        }
          int a2 = 1; 
        private void simpleButton3_Click(object sender, EventArgs e)//
            
        {
          
            if (a2 == 1)
            {
                layoutControlGroup4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                a2++;                
            }
            else
            {
                layoutControlGroup4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                a2 = 1;
            }
        
        }

        private void btnCapNhatChiTieu_Click(object sender, EventArgs e)//
        {
            if (layoutControlItem29.Text != "Đã xuất")
            {
                disEnable(false);
                if (layoutControlItem29.Text == "Đã chuyển" || layoutControlItem29.Text == "Hoàn thành")
                {
                    textEdit3.ReadOnly = true;
                    cboNoiKiem_18.ReadOnly = true;
                }
            }

            else
            {
                //XtraMessageBox.Show("Chỉ tiêu đã xuất không thể cập nhật thông tin");
               // disEnable(true);

                disEnable(false);

                
            }
               
        }

        private void btnXoaChiTieu_Click(object sender, EventArgs e)//
        {
            if (gridView44.RowCount > 1)
            {
                if (layoutControlItem29.Text != "Đã xuất")
                {
                    if (XtraMessageBox.Show("Bạn có chắc chắn xóa?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (cnmMod.XoaChiTieuKiem(layoutControlItem28.Text))
                            XtraMessageBox.Show("Xóa thành công");
                        else
                            XtraMessageBox.Show("Xóa thất bại");
                    }
                    load_CNCTK();
                }
                else
                    XtraMessageBox.Show("Chỉ tiêu đã xuất không thể xóa");
            }
            else
            {
                XtraMessageBox.Show("Không thể xóa hết chỉ tiêu kiểm");
            }
            
            
        }

        private void simpleButton5_Click(object sender, EventArgs e)//Button lưu thêm mới phương pháp
        {
            ppObj.Tenpp = textEdit4.Text;
            if (ppObj.Tenpp != "")
            {
                if (ctMod.KTPhuongPhap(ppObj) <= 0)
                {
                    if (ctMod.ThemPhuongPhap(ppObj))
                    {
                        XtraMessageBox.Show("Đã thêm phương pháp");
                    }
                }
                else { XtraMessageBox.Show("Phương pháp đã tồn tại"); }
                loadPP();
                simpleButton3_Click(sender, e);
                textEdit4.Text = ppObj.Tenpp;
            }
            else
                XtraMessageBox.Show("Chưa điền tên phương pháp");
        }

        private void simpleButton4_Click(object sender, EventArgs e)//Button cập nhật chỉ tiêu
        {

                
                gctObj.Maphieu = layoutControlItem28.Text;
                gctObj.Mact = layoutControlItem21.Text;

                gctObj.Mapp = textEdit3.EditValue.ToString();
                gctObj.Tenpp = textEdit3.Text;
                if (gctObj.Mapp == "")
                {
                    gctObj.Mapp = "NULL";
                    gctObj.Tenpp = "NULL";
                }
                gctObj.Mank = cboNoiKiem_18.EditValue.ToString();
                gctObj.Tennk = cboNoiKiem_18.Text;
                if (gctObj.Mank == "")
                {
                    gctObj.Mank = "NULL";
                    gctObj.Tennk = "NULL";
                }
                gctObj.Gia = txtDonGia.Text;
                if (cboLoaiGia.Text == "")
                    gctObj.Tenphanloaigia = gctObj.Gia;

                if (cnmMod.CapNhatChiTieuKiem(gctObj, txtMaMau.Text) == true)
                    if (cboLoaiGia.Text == "")
                        if (XtraMessageBox.Show("Giá này là một giá mới bạn có muốn thêm mới?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            cnmMod.AddGiaChiTieu(gctObj);
                            XtraMessageBox.Show("Đã thêm mới thành công");

                        }
                load_CNCTK();
            
          
            
        }

        private void btnYeuCauKiemLai_Click(object sender, EventArgs e)
        {
            if (textEdit6.Text != "")
            {
                if (XtraMessageBox.Show("Bạn chắc rằng yêu cầu kiểm lại", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    cnmMod.YeuCauKiemLai(layoutControlItem28.Text);
                    load_CNCTK();
                }
            }
            else
                XtraMessageBox.Show("Mẫu chưa kiểm");
        }

        private void layoutControlItem29_TextChanged(object sender, EventArgs e)
        {
            if (layoutControlItem29.Text == "Đã xuất")
            {
                //layoutControlItem19.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem16.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            }
            else
            {
               // layoutControlItem19.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem16.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }

        private void layoutControlItem21_TextChanged(object sender, EventArgs e)
        {
            cboLoaiGia.Properties.DataSource = cnmMod.GetLoaiGia(layoutControlItem21.Text);
            cboLoaiGia.Properties.ValueMember = "GIA";
            cboLoaiGia.Properties.DisplayMember = "TENPHANLOAIGIA";
        }

        private void cboChiDinh_18_EditValueChanged(object sender, EventArgs e)
        {
            if (cboChiDinh_18.Text != "" || cboChiDinh_18.EditValue != null)
            {
                int[] k = gridView1.GetSelectedRows();
                DataRow row = gridView1.GetDataRow(k[0]);
                macd_18 = row[0].ToString();
            }
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void simpleButton7_Click(object sender, EventArgs e)
        {
            load_CNCTK();
        }
    }
}