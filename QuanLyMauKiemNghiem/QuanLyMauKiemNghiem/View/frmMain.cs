using DevExpress.Utils.Menu;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Menu;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPrinting.Native.LayoutAdjustment;
using DevExpress.XtraReports.UI;
using DevExpress.XtraSplashScreen;
using Microsoft.Office.Interop.Excel;
using QuanLyMauKiemNghiem.Model;
using QuanLyMauKiemNghiem.Object;
using QuanLyMauKiemNghiem.Report;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyMauKiemNghiem.View
{
    public partial class frmMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        #region Khai Báo Các Biến

        private Chi_Tieu_Obj ctObj = new Chi_Tieu_Obj();
        private ChiTieu_Mod ctMod = new ChiTieu_Mod();
        private Nen_Mau_Obj nmObj = new Nen_Mau_Obj();
        private Don_Vi_Mod dvMod = new Don_Vi_Mod();
        private Quy_Chuan_Obj qcObj = new Quy_Chuan_Obj();
        private Quy_Chuan_Mod qcMod = new Quy_Chuan_Mod();
        private Don_Vi_Obj dvObj = new Don_Vi_Obj();
        private Phan_Loai_Gia_Obj plgObj = new Phan_Loai_Gia_Obj();
        private Khach_Hang_Mod khMod = new Khach_Hang_Mod();
        private Khach_Hang_Obj khObj = new Khach_Hang_Obj();
        private Phuong_Phap_Mod ppMod = new Phuong_Phap_Mod();
        private Phuong_Phap_Obj ppObj = new Phuong_Phap_Obj();
        private Noi_Kiem_Mod nkMod = new Noi_Kiem_Mod();
        private Noi_Kiem_Obj nkObj = new Noi_Kiem_Obj();
        private Nen_Mau_Mod nmMod = new Nen_Mau_Mod();
        private PhanLoaiGia_Mod plgMod = new PhanLoaiGia_Mod();
        private HopDong_Mod hdMod = new HopDong_Mod();
        private Hop_Dong_Obj hdObj = new Hop_Dong_Obj();
        private Nhan_Mau_Mod nhanmauMod = new Nhan_Mau_Mod();
        private Phieu_Nhan_Mau_Obj nhanmauObj = new Phieu_Nhan_Mau_Obj();
        private Phieu_YCPT_Obj pycptObj = new Phieu_YCPT_Obj();
        private Nhap_JobNo_Mod njnMod = new Nhap_JobNo_Mod();
        private CapNhatMau_Mod cnmMod = new CapNhatMau_Mod();
        private Cong_No_Mod cnMod = new Cong_No_Mod();
        private Chi_Dinh_Mod cdMod = new Chi_Dinh_Mod();
        private Cong_No_Obj cnObj = new Cong_No_Obj();
        private Chi_Dinh_Obj cdObj = new Chi_Dinh_Obj();
        private XuatYCPT_Mod xuat_ycpt_Mod = new XuatYCPT_Mod();
        private CongNoChuaKQ_Mod cnk_Mod = new CongNoChuaKQ_Mod();
        ThongKe_Mod tkMod = new ThongKe_Mod();


        private DuyetKQ_Mod dkqMod = new DuyetKQ_Mod();
        public static string mahd_cnm, tenmau_cnm, ngay_cnm, tempMAMAU_CNM;
        private LoadData l = new LoadData();
        private System.Data.DataTable dt1;
        private System.Data.DataTable dtJobNoVuaNhap = new System.Data.DataTable();
        private int themsua = 0;

        //Nhap ket qua
        private System.Data.DataTable dtNhapKQ;

        private QR_Mod qrMod = new QR_Mod();
        private NhapKetQua_Mod nkqMod = new NhapKetQua_Mod();

        //Xuất kết quả
        private XuatKQPT_Mod xkq = new XuatKQPT_Mod();

        private System.Data.DataTable dtXuatKQ;
        private System.Data.DataTable dtXuatKQ_PhuLuc;

        //Tạo QR Code
        private System.Data.DataTable dtTaoQR;

        //Ngươi dùng
        private Nguoi_dung_Mod ndMod = new Nguoi_dung_Mod();
        Nguoi_Dung_Obj ndObj = new Nguoi_Dung_Obj();
        SD_Quyen_Obj sdObj = new SD_Quyen_Obj();
        int themNguoiDung = 0;

        //Nhân bản JobNo
        Nhan_Ban_JobNo_Mod nbjnMod = new Nhan_Ban_JobNo_Mod();
        System.Data.DataTable dtNhanBan;

        //Nhân bản mẫu
        Nhan_Ban_Mau_Mod nbmMod = new Nhan_Ban_Mau_Mod();
        System.Data.DataTable dtNhanBan2;

        #endregion


        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            //Phân quyền
 	    if (ndMod.phanquyen(Login.tenUser, 5) > 0)//QUyền lập báo cáo thống kê
            {
                navBarGroup5.Visible = true;
                navigationFrame1.SelectedPage = navigationPage18;
                loadCongNo();
            }
           
            if (ndMod.phanquyen(Login.tenUser, 6) > 0)//Quyền nhân bản mẫu
            {
                navBarGroup6.Visible = true;
                navigationFrame1.SelectedPage = navigationPage21;
                loadNhanBanJOBNO();
            }
            
            if (ndMod.phanquyen(Login.tenUser, 1) > 0)//Quyền quản lý danh mục
            {
                navBarGroup1.Visible = true;
                navigationFrame1.SelectedPage = DSKhachHang;
                loadKH();
            }
            if (ndMod.phanquyen(Login.tenUser, 2) > 0)//Quyền Nhận mẫu trả kết quả
            {
                navBarGroup3.Visible = true;
                navigationFrame1.SelectedPage = navigationPage4;
                loadDStheodoiMau();
            }
            if (ndMod.phanquyen(Login.tenUser, 3) > 0)//Quyền Nhập kết quả
            {
                navBarGroup4.Visible = true;
                itemNhapKQ.Visible = true;
                navigationFrame1.SelectedPage = navigationPage13;
                loadDanhSachMauKQ();

            }
            if (ndMod.phanquyen(Login.tenUser, 4) > 0)//Quyền duyệt kết quả
            {
                navBarGroup4.Visible = true;
                itemNhapKQ.Visible = true;
                itemDuyetKQ.Visible = true;               
                itemDsTrangThaiKiemNghiem.Visible = true;
                navigationFrame1.SelectedPage = navigationPage14;
                Load_Form_Duyet_KQ();
            }
           

            if (ndMod.phanquyen(Login.tenUser, 7) > 0)//Quyền quản lý người dùng
            {
                barButtonItem9.Visibility = DevExpress.XtraBars.BarItemVisibility.OnlyInRuntime;
                navigationFrame1.SelectedPage = navigationPage20;
                loadNguoiDung();

            }
            //

            skin();
            //loadKH();
            //loadPhuongPhap();
            //loadChiDinh();
            //Load_Chi_Tieu();
            //loadDanhSachMauKQ();//Load kết quả
            //Load_Form_Duyet_KQ();//Load duyệt kết quả
            //loadXuatKetQua();//Load xuất kết quả trả khách hàng
            //loadQRCode();//Load tạo QR code
            //loadNguoiDung();//Load người dùng
            //loadNhanBanJOBNO();//Load nhân bản theo jobno
            //loadNhanMau();//Load hợp đồng trong nhận mẫu
            //Load_Form_Nhap_JobNo();//Load cập nhật jobno
            //loadNhanBanMau();// Load nhân bản theo mẫu
            //loadCapNhatMau();//Load cập nhật thông tin mẫu
            loadCongNo(); //Load Công nợ
            //loadCongNoK();//Load cong nợ chưa xuất
            //loadXuatPhieu_YCPT();//Load xuất phiếu yêu cầu phân tích            
            //loadNenMau();//Load nền mẫu
            //Load_Hop_Dong();//Load hợp đồng
            //loadQuyChuan();
            //loadDonVi();
            //loadNoiKiem();
            //loadDStheodoiMau();//Load danh sách theo dõi mẫu
            //loadDStheodoikiemNghiem();//load danh sách theo dõi của bộ phận kiểm nghiệm
            //loadThongKeChiTieu();//Thống kê chỉ tiêu kiểm

            //NHân mẫu
            tableNhanMau();

            dtJobNoVuaNhap.Columns.Add("JOBNO");
            dtJobNoVuaNhap.Columns.Add("MAMAU");
            dtJobNoVuaNhap.Columns.Add("TENMAU");
            dtJobNoVuaNhap.Columns.Add("SLMAU");
            dtJobNoVuaNhap.Columns.Add("KHOILUONGMAU");
            dtJobNoVuaNhap.Columns.Add("TIME_NHANMAU_DK");
            dtJobNoVuaNhap.Columns.Add("TIME_NHANMAU_TT");
            dtJobNoVuaNhap.Columns.Add("temp_MAMAU");
            //Nhân bản JOBNO
            dtNhanBan = new System.Data.DataTable();
            dtNhanBan.Columns.Add("TENKH");
            dtNhanBan.Columns.Add("MAHD");
            dtNhanBan.Columns.Add("JOBNO");
            dtNhanBan.Columns.Add("MAMAU");
            dtNhanBan.Columns.Add("TENMAU");
            dtNhanBan.Columns.Add("TIME_NHANMAU_DK");
            dtNhanBan.Columns.Add("MOTAMAU");
            dtNhanBan.Columns.Add("NGUONMAU");
            dtNhanBan.Columns.Add("KHOILUONGMAU");
            dtNhanBan.Columns.Add("TINHTRANGMAU");
            dtNhanBan.Columns.Add("NGUOINHANMAU");
            dtNhanBan.Columns.Add("SEAL");
            dtNhanBan.Columns.Add("TIME_TRABCPT_DK");
            dtNhanBan.Columns.Add("TTKHCUNGCAP");
            dtNhanBan.Columns.Add("temp_MAMAU");

            //Nhân bản THEO MẪU
            dtNhanBan2 = new System.Data.DataTable();
            dtNhanBan2.Columns.Add("TENKH");
            dtNhanBan2.Columns.Add("MAHD");
            dtNhanBan2.Columns.Add("JOBNO");
            dtNhanBan2.Columns.Add("MAMAU");
            dtNhanBan2.Columns.Add("TENMAU");
            dtNhanBan2.Columns.Add("TIME_NHANMAU_DK");
            dtNhanBan2.Columns.Add("MOTAMAU");
            dtNhanBan2.Columns.Add("NGUONMAU");
            dtNhanBan2.Columns.Add("KHOILUONGMAU");
            dtNhanBan2.Columns.Add("TINHTRANGMAU");
            dtNhanBan2.Columns.Add("NGUOINHANMAU");
            dtNhanBan2.Columns.Add("SEAL");
            dtNhanBan2.Columns.Add("TIME_TRABCPT_DK");
            dtNhanBan2.Columns.Add("TTKHCUNGCAP");
            dtNhanBan2.Columns.Add("temp_MAMAU");
        }
        

        #region DANH MỤC

        //Button load lại dữ liệu
        private void barButtonItem23_ItemClick(object sender, ItemClickEventArgs e)
        {
            loadKH();
            loadPhuongPhap();
            loadChiDinh();
            Load_Chi_Tieu();
            loadDanhSachMauKQ();//Load kết quả
            Load_Form_Duyet_KQ();//Load duyệt kết quả
            loadXuatKetQua();//Load xuất kết quả trả khách hàng
            loadQRCode();//Load tạo QR code
            loadNguoiDung();//Load người dùng
            loadNhanBanJOBNO();//Load nhân bản theo jobno
            loadNhanMau();//Load hợp đồng trong nhận mẫu
            Load_Form_Nhap_JobNo();//Load cập nhật jobno
            loadNhanBanMau();// Load nhân bản theo mẫu
            loadCapNhatMau();//Load cập nhật thông tin mẫu
            loadCongNo(); //Load Công nợ
            loadCongNoK();//Load cong nợ chưa xuất
            loadXuatPhieu_YCPT();//Load xuất phiếu yêu cầu phân tích            
            loadNenMau();//Load nền mẫu
            Load_Hop_Dong();//Load hợp đồng
            loadQuyChuan();
            loadDonVi();
            loadNoiKiem();
            loadDStheodoiMau();//Load danh sách theo dõi mẫu
            loadDStheodoikiemNghiem();//load danh sách theo dõi của bộ phận kiểm nghiệm
            loadThongKeChiTieu();//Thống kê chỉ tiêu kiểm
        }

        private void barButtonItem24_ItemClick(object sender, ItemClickEventArgs e)//Đăng xuất
        {
            this.Hide();
            Login f = new Login();
            f.ShowDialog();
        }

        private void itemNhanBanMau_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage22;
            loadNhanBanMau();
        }

        private void navBarItem4_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage21;
            loadNhanBanJOBNO();

        }

        private void barButtonItem12_ItemClick(object sender, ItemClickEventArgs e)
        {
            frmDoiMatKhau fm = new frmDoiMatKhau();
            fm.ShowDialog();
        }
       public static int connect = 0;
        private void barButtonItem13_ItemClick(object sender, ItemClickEventArgs e)
        {
            frmConnectDB fm = new frmConnectDB();
            this.Hide();
            fm.FormClosed+=new FormClosedEventHandler(frmConnectDB_FormClosed);
            fm.ShowDialog();
            //if (connect == 1)
            //{
            //    this.Hide();
            //    Login f = new Login();
            //    f.ShowDialog();
            //}              
           
        }
        private void frmConnectDB_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Show();
        }

        private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage20;
            loadNguoiDung();
        }

        
        private void navBarItem8_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//Xuất QR Code
        {
            navigationFrame1.SelectedPage = navigationPage19;
            loadQRCode();
        }
        private void itemTKCN_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage18;
            loadCongNo();
            loadCongNoK();
        }

        private void itemXuatPKQ_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage16;
            loadXuatKetQua();
        }

        private void barButtonItem14_ItemClick(object sender, ItemClickEventArgs e)
        {
            this.pdfViewer1.LoadDocument(@".\HDSD.pdf");
            navigationFrame1.SelectedPage = navigationPage15;
        }

        private void itemDuyetKQ_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage14;
            Load_Form_Duyet_KQ();
        }

        private void itemNhapKQ_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage13;
            loadDanhSachMauKQ();
        }

        private void itemKH_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = DSKhachHang;
            loadKH();
        }

        private void itemCNJobNo_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage12;
            Load_Form_Nhap_JobNo();
            /////
        }

        private void itemChiTieu_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage2;
            Load_Chi_Tieu();
        }

        private void ItemCNTTMau_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage11;
            loadCapNhatMau();
        }

        private void skin()
        {
            DevExpress.LookAndFeel.DefaultLookAndFeel giaodien = new DevExpress.LookAndFeel.DefaultLookAndFeel();
            giaodien.LookAndFeel.SkinName = "Office 2007 Blue";
        }

        private void navBarItem16_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage17;
            loadXuatPhieu_YCPT();

        }

        private void itemPP_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage3;
            loadPhuongPhap();
        }

        private void navBarItem6_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage5;
            loadChiDinh();
        }

        private void itemHopDong_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage9;
            Load_Hop_Dong();
        }

        private void itemNenMau_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage7;
            loadNenMau();
        }



        private void ItemQC_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)//
        {
            navigationFrame1.SelectedPage = navigationPage8;
            loadQuyChuan();
        }

        private void navBarItem5_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage23;
            loadThongKeChiTieu();

        }

        private void itemDsTrangThaiKiemNghiem_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage24;
            loadDStheodoikiemNghiem();
        }

        private void navBarItem7_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage4;
            loadDStheodoiMau();
        }


        #endregion Danh Mục 

        #region QUẢN LÝ PHƯƠNG PHÁP

        private void simpleButton33_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls|Excel Workbook|*.xlsm", ValidateNames = true };
            textBox5.Text = ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : "";
        }

        private void simpleButton72_Click(object sender, EventArgs e)
        {
            if (textBox5.Text == "" || textBox5.Text == "Chưa có tệp nào được chọn")
            {
                XtraMessageBox.Show("Vui lòng nhập đường dẫn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox5.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] a = new string[colCount];
            int sl = 0, sl2 = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 == null)
                    {
                        a[j - 2] = "";
                    }
                    else
                    {
                        a[j - 2] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                ppObj.Tenpp = a[0];

                if (!ppMod.TrungTenPP(ppObj))
                {
                    ppMod.addData(ppObj);
                    sl++;
                }
                else
                {
                    sl2++;
                }
            }
            XtraMessageBox.Show("Đã thêm " + sl + " phương pháp vào CSDL!  --Có " + sl2 + " phương pháp đã tồn tại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.None);
            loadPhuongPhap();
            textBox5.Text = "Chưa có tệp nào được chọn";
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }

        private void simpleButton73_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    XtraMessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "Phương Pháp";
                if (ws == null)
                {
                    XtraMessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                int fontSizeTenTruong = 14;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A1");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 5.5;

                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B1");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN PHƯƠNG PHÁP";
                row23_MaSP.ColumnWidth = 50;

                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A1", "B1");

                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                //in đậm
                row23_CotTieuDe.Font.Bold = true;

                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);

                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A1", "B1"));

                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        // Nút Sửa PP
        void enablePhuongphap(bool e)
        {
            txtTenPhuongPhap.ReadOnly = e;
            dgvPhuongPhap.Enabled = e;
            if (e)
            {
                txtTenPhuongPhap.BackColor = Color.Linen;
                layoutControlItem114.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem91.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }

            else
            {
                txtTenPhuongPhap.BackColor = Color.White;

                layoutControlItem114.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem91.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            }

        }

        private void capnhatPhuongPhap_Click(object sender, EventArgs e)
        {
            enablePhuongphap(false);
        }

        private void xoaphuongphap_Click(object sender, EventArgs e)
        {
            string mapp = txtMaPP.Text;

            DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (ppMod.delectData(mapp))
                    XtraMessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    XtraMessageBox.Show("Xóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                return;
            loadPhuongPhap();
        }


        private void loadPhuongPhap() //Load danh sách
        {
            themsua = 0;
            enablePhuongphap(true);
            dgvPhuongPhap.DataSource = ppMod.GetDataPP();

            txtMaPP.DataBindings.Clear();
            txtMaPP.DataBindings.Add("Text", dgvPhuongPhap.DataSource, "MAPP");
            txtTenPhuongPhap.DataBindings.Clear();
            txtTenPhuongPhap.DataBindings.Add("Text", dgvPhuongPhap.DataSource, "TENPP");
            ErrorPP.Text = "";

        }



        private void simpleButton18_Click(object sender, EventArgs e) //thêm mới phương pháp
        {
            themsua = 3;

            txtMaPP.Text = "";
            txtTenPhuongPhap.Text = "";
            enablePhuongphap(false);
        }

        private void simpleButton16_Click(object sender, EventArgs e) //Load lại danh sách phương pháp
        {
            loadPhuongPhap();
        }

        private void simpleButton17_Click(object sender, EventArgs e)   // LƯU
        {
            ppObj.Tenpp = txtTenPhuongPhap.Text;
            ppObj.Mapp = txtMaPP.Text;
            if (txtTenPhuongPhap.Text == "")
            {
                ErrorPP.Text = "Vui lòng nhập tên phương pháp";
                return;
            }
            else
            {
                if (themsua == 3)
                {
                    if (ppMod.kiemTraTenPP(ppObj.Tenpp).Rows.Count > 0)
                    {
                        ErrorPP.Text = "Phương pháp đã tồn tại!";
                        return;
                    }
                    if (ppMod.addData(ppObj))
                    {
                        if (XtraMessageBox.Show("Đã thêm phương pháp, bạn có muốn tiếp tục thêm?", "Thông báo", MessageBoxButtons.YesNo) != DialogResult.No)
                        {
                            simpleButton18_Click(sender, e);
                            return;
                        }

                        // MessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //loadPhuongPhap();
                    }
                    else
                    {
                        XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    if (ppMod.update(ppObj))
                    {
                        XtraMessageBox.Show("Cập nhật thông tin thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                loadPhuongPhap();
            }
        }

        #endregion Phương Pháp

        #region QUẢN LÝ DANH MỤC KHÁCH HÀNG
        void enableKhachHang(bool e)
        {
            dgvKhachHang.Enabled = e;
            txtTenKH.ReadOnly = e;
            txtDiaChiKH.ReadOnly = e;
            txtEmailKH.ReadOnly = e;
            txtSDT.ReadOnly = e;
            txtSoFax.ReadOnly = e;
            txtMST.ReadOnly = e;
            txtNguoiLH.ReadOnly = e;
            txtEmailLH.ReadOnly = e;
            txtSDTLH.ReadOnly = e;
            if (e)
            {
                txtTenKH.BackColor = Color.Linen;
                txtDiaChiKH.BackColor = Color.Linen;
                txtEmailKH.BackColor = Color.Linen;
                txtSDT.BackColor = Color.Linen;
                txtSoFax.BackColor = Color.Linen;
                txtMST.BackColor = Color.Linen;
                txtNguoiLH.BackColor = Color.Linen;
                txtEmailLH.BackColor = Color.Linen;
                txtSDTLH.BackColor = Color.Linen;
                layoutControlItem260.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem284.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }

            else
            {

                txtTenKH.BackColor = Color.White;
                txtDiaChiKH.BackColor = Color.White;
                txtEmailKH.BackColor = Color.White;
                txtSDT.BackColor = Color.White;
                txtSoFax.BackColor = Color.White;
                txtMST.BackColor = Color.White;
                txtNguoiLH.BackColor = Color.White;
                txtEmailLH.BackColor = Color.White;
                txtSDTLH.BackColor = Color.White;
                layoutControlItem260.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem284.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            }

        }

        // Nút Sửa KH        

        private void capnhatKhachHang_Click(object sender, EventArgs e)
        {
            themsua = 1;
            enableKhachHang(false);
        }
        private void Xóa_Click(object sender, EventArgs e) //xóa khách hàng
        {
            string makh = txtMaKH.Text;
            string count = khMod.check_Xoa_kh(makh).Rows[0]["MAKH"].ToString();
            if (count == "0")
            {
                DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    if (khMod.delectData(txtMaKH.Text))
                        XtraMessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        XtraMessageBox.Show("Xóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                    return;
                loadKH();
            }
            else
            {
                XtraMessageBox.Show("Bạn không thể xóa khách hàng này");
            }
        }


        //Load
        public void loadKH()
        {
            themsua = 1;
            enableKhachHang(true);

            dgvKhachHang.DataSource = khMod.GetDataKH();
            txtMaKH.DataBindings.Clear();
            txtMaKH.DataBindings.Add("Text", dgvKhachHang.DataSource, "MAKH");
            txtEmailKH.DataBindings.Clear();
            txtEmailKH.DataBindings.Add("Text", dgvKhachHang.DataSource, "EMAIL");
            txtSDTLH.DataBindings.Clear();
            txtSDTLH.DataBindings.Add("Text", dgvKhachHang.DataSource, "SDT_LH");
            txtNguoiLH.DataBindings.Clear();
            txtNguoiLH.DataBindings.Add("Text", dgvKhachHang.DataSource, "NGUOI_LH");

            txtMST.DataBindings.Clear();
            txtMST.DataBindings.Add("Text", dgvKhachHang.DataSource, "MASOTHUE");
            txtSDT.DataBindings.Clear();
            txtSDT.DataBindings.Add("Text", dgvKhachHang.DataSource, "SDT");
            txtTenKH.DataBindings.Clear();
            txtTenKH.DataBindings.Add("Text", dgvKhachHang.DataSource, "TENKH");
            txtDiaChiKH.DataBindings.Clear();
            txtDiaChiKH.DataBindings.Add("Text", dgvKhachHang.DataSource, "DIACHI");
            txtSoFax.DataBindings.Clear();
            txtSoFax.DataBindings.Add("Text", dgvKhachHang.DataSource, "SOFAX");

            txtEmailLH.DataBindings.Clear();
            txtEmailLH.DataBindings.Add("Text", dgvKhachHang.DataSource, "EMAIL_LH");
            txtFilePath.Text = "Chưa có tệp nào được chọn.";
        }

        private void simpleButton9_Click(object sender, EventArgs e)//Button load lại
        {
            loadKH();
        }


        private void btnThemFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls|Excel Workbook|*.xlsm", ValidateNames = true };

            txtFilePath.Text = ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : "";
        }


        private void btnThemMoi_Click(object sender, EventArgs e)
        {
            themsua = 2;
            enableKhachHang(false);

            txtSDT.Text = "";
            txtTenKH.Text = "";
            txtDiaChiKH.Text = "";
            txtSoFax.Text = "";
            txtEmailKH.Text = "";
            txtMaKH.Text = "";
            txtMaKH.Text = l.AutoID("KH-", "SELECT MAX(RIGHT(MAKH,7)) from KHACH_HANG");
            txtEmailKH.Text = "";
            txtNguoiLH.Text = "";
            txtNguoiLH.Text = "";
            txtMST.Text = "";
        }  //Button thêm mới 

        private void simpleButton53_Click(object sender, EventArgs e) //Button lưu khách hàng
        {
            SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
            //The Wait Form is opened in a separate thread. To change its Description, use the SetWaitFormDescription method.
            for (int i = 90; i <= 100; i++)
            {
                SplashScreenManager.Default.SetWaitFormDescription(i.ToString() + "%");
                Thread.Sleep(1);
            }

            //Close Wait Form
            SplashScreenManager.CloseForm(false);

            khObj.Makh = txtMaKH.Text;
            khObj.Tenkh = txtTenKH.Text;
            khObj.Diachi = txtDiaChiKH.Text;
            khObj.Masothue = txtMST.Text;
            khObj.Sdt = txtSDT.Text;
            khObj.Email = txtEmailKH.Text;
            khObj.Sofax = txtSoFax.Text;
            khObj.Nguoi_lh = txtNguoiLH.Text;
            khObj.Sdt_lh = txtNguoiLH.Text;
            khObj.Email_lh = txtEmailLH.Text;

            if (txtTenKH.Text == "" || txtDiaChiKH.Text == "" || txtSDT.Text == "" || txtEmailKH.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ thông tin có dấu *");
                return;
            }
            else
            {
                if (themsua == 2)
                {
                    if (khMod.addData(khObj))
                    {
                        XtraMessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //    frmMain_Load(sender, e);
                        loadKH();
                    }
                    else
                    {
                        XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    if (khMod.update(khObj))
                    {
                        XtraMessageBox.Show("Cập nhật thông tin thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadKH();
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e) // Thêm khách hàng từ excel(button Hoàn thành)
        {
            if (txtFilePath.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đường dẫn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("TENKH", typeof(string));
            dt.Columns.Add("DIACHI", typeof(string));
            dt.Columns.Add("MASOTHUE", typeof(string));
            dt.Columns.Add("SDT", typeof(string));
            dt.Columns.Add("EMAIL", typeof(string));
            dt.Columns.Add("FAX", typeof(string));
            dt.Columns.Add("NGUOI_LH", typeof(string));
            dt.Columns.Add("SDT_LH", typeof(string));
            dt.Columns.Add("EMAIL_LH", typeof(string));

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtFilePath.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] a = new string[colCount];
            int test = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 == null)
                    {
                        a[j - 2] = "";
                    }
                    else
                    {
                        a[j - 2] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                dt.Rows.Add(a[0], a[1], a[4], a[2], a[3], a[5], a[6], a[7], a[8]);
                txtMaKH.Text = l.AutoID("KH-", "SELECT MAX(RIGHT(MAKH,7)) from KHACH_HANG");
                khObj.Makh = txtMaKH.Text;
                khObj.Tenkh = a[0];
                khObj.Diachi = a[1];
                khObj.Masothue = a[4];
                khObj.Sdt = a[2];
                khObj.Email = a[3];
                khObj.Sofax = a[5];
                khObj.Nguoi_lh = a[6];
                khObj.Sdt_lh = a[7];
                khObj.Email_lh = a[8];
                if (!khMod.addData(khObj))
                {
                    XtraMessageBox.Show("Thêm KH thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else test++;
            }
            dgvKhachHang.DataSource = dt;
            XtraMessageBox.Show("Đã thêm " + test + "dòng vào CSDL!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.None);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            loadKH();
        }


        private void btn_Click(object sender, EventArgs e) //Tải file Excel (button tải mẫu file Excel)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    XtraMessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Workbook wb = xlApp.Workbooks.Add(misValue);

                Worksheet ws = (Worksheet)wb.Worksheets[1];
                ws.Name = "Khach Hang";
                if (ws == null)
                {
                    XtraMessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                // int fontSizeTieuDe = 18;
                int fontSizeTenTruong = 14;
                // int fontSizeNoiDung = 12;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A1");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 5.5;

                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B1");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN KHÁCH HÀNG";
                row23_MaSP.ColumnWidth = 50;

                //Cột 3: Tạo Ô Mã Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaS = ws.get_Range("C1");//Cột B dòng 2 và dòng 3
                row23_MaS.Merge();
                row23_MaS.Font.Size = fontSizeTenTruong;
                row23_MaS.Font.Name = fontName;
                row23_MaS.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_MaS.Value2 = "ĐỊA CHỈ";
                row23_MaS.ColumnWidth = 50;

                // Cột 4: Tạo Ô Tên Mẫu:
                Microsoft.Office.Interop.Excel.Range row23_TenSP = ws.get_Range("D1");//Cột C dòng 2 và dòng 3
                row23_TenSP.Merge();
                row23_TenSP.Font.Size = fontSizeTenTruong;
                row23_TenSP.Font.Name = fontName;
                row23_TenSP.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_TenSP.ColumnWidth = 20;
                row23_TenSP.Value2 = "SDT";

                //Cột 5: Tạo các ô  chỉ tiêu :
                Microsoft.Office.Interop.Excel.Range row2_GiaSP = ws.get_Range("E1");//Cột D->E của dòng 2
                row2_GiaSP.Merge();
                row2_GiaSP.Font.Size = fontSizeTenTruong;
                row2_GiaSP.Font.Name = fontName;
                row2_GiaSP.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row2_GiaSP.ColumnWidth = 50;
                row2_GiaSP.Value2 = "EMAIL";

                //Cột 6: Tạo các ô  chỉ tiêu :
                Microsoft.Office.Interop.Excel.Range row23_NenMau = ws.get_Range("F1");//Cột D->E của dòng 2
                row23_NenMau.Merge();
                row23_NenMau.Font.Size = fontSizeTenTruong;
                row23_NenMau.Font.Name = fontName;
                row23_NenMau.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_NenMau.ColumnWidth = 20;
                row23_NenMau.Value2 = "MÃ SỐ THUẾ";

                //Cột 7: Tạo các ô  chỉ tiêu :
                Microsoft.Office.Interop.Excel.Range row23_SoFax = ws.get_Range("G1");//Cột D->E của dòng 2
                row23_SoFax.Merge();
                row23_SoFax.Font.Size = fontSizeTenTruong;
                row23_SoFax.Font.Name = fontName;
                row23_SoFax.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_SoFax.ColumnWidth = 20;
                row23_SoFax.Value2 = "SỐ FAX";

                //Cột 8: Tạo các ô  chỉ tiêu :
                Microsoft.Office.Interop.Excel.Range row23_Nguoi_LH = ws.get_Range("H1");//Cột D->E của dòng 2
                row23_Nguoi_LH.Merge();
                row23_Nguoi_LH.Font.Size = fontSizeTenTruong;
                row23_Nguoi_LH.Font.Name = fontName;
                row23_Nguoi_LH.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_Nguoi_LH.ColumnWidth = 20;
                row23_Nguoi_LH.Value2 = "NGƯỜI LIÊN HỆ";

                //Cột 9: Tạo các ô  chỉ tiêu :
                Microsoft.Office.Interop.Excel.Range row23_SDT_LH = ws.get_Range("I1");//Cột D->E của dòng 2
                row23_SDT_LH.Merge();
                row23_SDT_LH.Font.Size = fontSizeTenTruong;
                row23_SDT_LH.Font.Name = fontName;
                row23_SDT_LH.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_SDT_LH.ColumnWidth = 20;
                row23_SDT_LH.Value2 = "SĐT NGƯỜI LIÊN HỆ";

                //Cột 9: Tạo các ô  chỉ tiêu :
                Microsoft.Office.Interop.Excel.Range row23_Email_LH = ws.get_Range("J1");//Cột D->E của dòng 2
                row23_Email_LH.Merge();
                row23_Email_LH.Font.Size = fontSizeTenTruong;
                row23_Email_LH.Font.Name = fontName;
                row23_Email_LH.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                row23_Email_LH.ColumnWidth = 20;
                row23_Email_LH.Value2 = "Email NGƯỜI LIÊN HỆ";

                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A1", "J1");

                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                //in đậm
                row23_CotTieuDe.Font.Bold = true;

                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);

                //int stt = 0;
                //row = 2;//dữ liệu xuất bắt đầu từ dòng số 4 trong file Excel (khai báo 3 để vào vòng lặp nó ++ thành 4)

                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A1", "J1"));

                //Lưu file excel xuống Ổ cứng
                //string saveExcelFile = @"e:\excel_report.xlsx";
                //wb.SaveAs(filePath);

                //đóng file để hoàn tất quá trình lưu trữ
                //wb.Close(true, misValue, misValue);

                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                //xlApp.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        //Hàm kẻ khung cho Excel
        private void BorderAround(Microsoft.Office.Interop.Excel.Range range)
        {
            Borders borders = range.Borders;
            borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            borders.Color = Color.Black;
            borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
            borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
        }

        //Hàm thu hồi bộ nhớ cho COM Excel
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                obj = null;
            }
            finally
            { GC.Collect(); }
        }

        private System.Data.DataTable ReadDataFromExcelFile()
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=Excel 12.0;";
            //string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=Excel 8.0";
            //string connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0 Xml;Data Source={0}; Extended Properties='Excel 12.0; HDR=Yes'", txtFilePath.Text.Trim());
            // Tạo đối tượng kết nối
            OleDbConnection oledbConn = new OleDbConnection(connectionString);
            System.Data.DataTable data = null;
            try
            {
                // Mở kết nối
                oledbConn.Open();

                // Tạo đối tượng OleDBCommand và query data từ sheet có tên "Sheet1"
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);

                // Tạo đối tượng OleDbDataAdapter để thực thi việc query lấy dữ liệu từ tập tin excel
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                // Tạo đối tượng DataSet để hứng dữ liệu từ tập tin excel
                DataSet ds = new DataSet();

                // Đổ đữ liệu từ tập excel vào DataSet
                oleda.Fill(ds);

                data = ds.Tables[0];
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
            finally
            {
                // Đóng chuỗi kết nối
                oledbConn.Close();
            }
            return data;
        }

        #endregion QUẢN LÝ DANH MỤC KHÁCH HÀNG
            

        #region QUẢN LÝ DANH MỤC CHỈ TIÊU
        // Nút Sửa CT
        void enableChiTieu(bool e)
        {
            gridControl2.Enabled = e;
            txtTenCT.ReadOnly = e;
            textEdit47.ReadOnly = e;
            cbNM.ReadOnly = e;
            cbDV.ReadOnly = e;
            cbPP.ReadOnly = e;
            cbQC.ReadOnly = e;
            cbPLG.ReadOnly = e;
            cbNM.ReadOnly = e;
            txtPL.ReadOnly = e;
            txtLOD.ReadOnly = e;
            txtGia.ReadOnly = e;

            if (e)
            {
                txtTenCT.BackColor = Color.Linen;
                textEdit47.BackColor = Color.Linen;
                cbNM.BackColor = Color.Linen;
                cbDV.BackColor = Color.Linen;
                cbPP.BackColor = Color.Linen;
                cbQC.BackColor = Color.Linen;
                cbPLG.BackColor = Color.Linen;
                cbNM.BackColor = Color.Linen;
                txtPL.BackColor = Color.Linen;
                txtLOD.BackColor = Color.Linen;
                txtGia.BackColor = Color.Linen;
                layoutControlItem286.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem287.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }

            else
            {

                txtTenCT.BackColor = Color.White;
                textEdit47.BackColor = Color.White;
                cbNM.BackColor = Color.White;
                cbDV.BackColor = Color.White;
                cbPP.BackColor = Color.White;
                cbQC.BackColor = Color.White;
                cbPLG.BackColor = Color.White;
                cbNM.BackColor = Color.White;
                txtPL.BackColor = Color.White;
                txtLOD.BackColor = Color.White;
                txtGia.BackColor = Color.White;
                layoutControlItem286.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem287.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            }

        }
        private void btnSuaChiTieu_Click(object sender, EventArgs e)
        {
            enableChiTieu(false);
        }

        private void btnXoaChiTieu_Click(object sender, EventArgs e)//Xóa chỉ tiêu
        {
            ctObj.Mact = txtMaCT.Text;
            DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (ctMod.kiemtratontaict(ctObj) != 0)
                {
                   if(ctMod.kiemtrasoluonggia(ctObj)<=1)
                   {
                        XtraMessageBox.Show("Chỉ tiêu này không thể xóa!");
                    return;
                   }
                }
                
                 if (ctMod.XoaChiTieu(ctObj, layoutControlItem307.Text))
                    {
                        XtraMessageBox.Show("Xóa chỉ tiêu thành công!");
                        Load_Chi_Tieu();
                    }
                    else { XtraMessageBox.Show("Xóa chỉ tiêu thất bại!"); return; }
                
                
            }
        }



        private int them; // them=1 (update) -- them=2(insert)

        // LOAD FORM
        private void Load_Chi_Tieu()
        {
            enableChiTieu(true);

            them = 1;
            layoutControlItem307.Text = "";
            System.Data.DataTable dt = ctMod.GetChiTieu();
            cbNM.Properties.DataSource = ctMod.GetNenMau();
            cbNM.Properties.ValueMember = "MANENMAU";
            cbNM.Properties.DisplayMember = "TENNENMAU";
            cbPP.Properties.DataSource = ctMod.GetPhuongPhap();
            cbPP.Properties.ValueMember = "MAPP";
            cbPP.Properties.DisplayMember = "TENPP";
            cbQC.Properties.DataSource = ctMod.GetQuyChuan();
            cbQC.Properties.ValueMember = "MAQC";
            cbQC.Properties.DisplayMember = "TENQC";
            cbDV.Properties.DataSource = ctMod.GetDonVi();
            cbDV.Properties.ValueMember = "MADV";
            cbDV.Properties.DisplayMember = "TENDV";
            cbPLG.Properties.DataSource = ctMod.GetPhanLoaiGia();
            cbPLG.Properties.ValueMember = "TENPHANLOAIGIA";
            cbPLG.Properties.DisplayMember = "TENPHANLOAIGIA";
            textEdit47.Properties.Items.Clear();
            System.Data.DataTable dtnhomct = ctMod.GetNhomChiTieu();
            foreach (DataRow row in dtnhomct.Rows)
            {
                textEdit47.Properties.Items.Add(row["NHOMCHITIEU"]);
            }

            //int a = dt.Rows.Count;
            gridControl2.DataSource = dt;
            txtMaCT.DataBindings.Clear();
            txtMaCT.DataBindings.Add("Text", gridControl2.DataSource, "MACT");
            txtTenCT.DataBindings.Clear();
            txtTenCT.DataBindings.Add("Text", gridControl2.DataSource, "TENCT");
            txtLOD.DataBindings.Clear();
            txtLOD.DataBindings.Add("Text", gridControl2.DataSource, "LOD");
            txtGia.DataBindings.Clear();
            txtGia.DataBindings.Add("Text", gridControl2.DataSource, "GIA");
            txtPL.DataBindings.Clear();
            txtPL.DataBindings.Add("Text", gridControl2.DataSource, "PHAN_LOAI");

            cbNM.DataBindings.Clear();
            cbNM.DataBindings.Add("EditValue", gridControl2.DataSource, "MANENMAU");
            cbPP.DataBindings.Clear();
            cbPP.DataBindings.Add("EditValue", gridControl2.DataSource, "MAPP");
            cbQC.DataBindings.Clear();
            cbQC.DataBindings.Add("EditValue", gridControl2.DataSource, "MAQC");
            cbDV.DataBindings.Clear();
            cbDV.DataBindings.Add("EditValue", gridControl2.DataSource, "MADV");
            cbPLG.DataBindings.Clear();
            cbPLG.DataBindings.Add("EditValue", gridControl2.DataSource, "TENPHANLOAIGIA");
            textEdit47.DataBindings.Clear();
            textEdit47.DataBindings.Add("EditValue", gridControl2.DataSource, "NHOMCHITIEU");
            layoutControlItem307.DataBindings.Clear();
            layoutControlItem307.DataBindings.Add("Text", gridControl2.DataSource, "MAGIA");
        }

        // NUT +
        private int a1 = 1, a2 = 1, a3 = 1, a4 = 1;

        //Nut Thêm
        private void simpleButton6_Click(object sender, EventArgs e)
        {
            them = 2;
            enableChiTieu(false);

            txtMaCT.Text = ctMod.autoMaChiTieu();
            txtTenCT.Text = string.Empty;
            txtLOD.Text = string.Empty;
            txtGia.Text = string.Empty;
            txtPL.Text = string.Empty;

            cbNM.Text = string.Empty;
            cbPP.Text = string.Empty;
            cbQC.Text = string.Empty;
            cbDV.Text = string.Empty;
            cbPLG.Text = string.Empty;
        }

        //Nút Lưu
        private void simpleButton7_Click(object sender, EventArgs e)
        {
            if (txtTenCT.Text == "" || cbNM.Text == "" || cbPP.Text == "" || cbDV.Text == "" || txtLOD.Text == "" || txtPL.Text == "" || txtGia.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập các thông tin có dấu (*)");
                return;
            }
            ctObj.Mact = txtMaCT.Text;
            ctObj.Tenct = txtTenCT.Text;
            ctObj.Manenmau = cbNM.EditValue.ToString();
            ctObj.Tennenmau = cbNM.Text;
            ctObj.Mapp = cbPP.EditValue.ToString();
            ctObj.Tenpp = cbPP.Text;
            ctObj.Maqc = cbQC.EditValue.ToString();
            ctObj.Tenqc = cbQC.Text;
            ctObj.Madv = cbDV.EditValue.ToString();
            ctObj.Tendv = cbDV.Text;
            ctObj.Lod = txtLOD.Text;
            ctObj.Phan_loai = txtPL.Text;
            ctObj.Nhom_ct = textEdit47.Text;
            ctObj.Gia = txtGia.Text;

            ctObj.Loaigia = cbPLG.Text;
            if (ctObj.Loaigia == "")
                ctObj.Loaigia = ctObj.Gia;

            if (them == 1)
            {
                try
                {
                    if (ctMod.kiemtragiachitieu(ctObj.Mact, ctObj.Loaigia).Rows.Count == 0)
                    {                            //them mới giá
                        ctMod.themMoiGiaChiTieu(ctObj, true);
                    }
                    else
                    {
                        ctMod.themMoiGiaChiTieu(ctObj, false);
                    }
                    if (!ctMod.UpdateChiTieu(ctObj))
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!");
                        return;
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thành công!");
                    }
                }
                catch (Exception ex) { XtraMessageBox.Show("" + ex); }
            }
            else
            {
                if (them == 2)
                {
                    if (!ctMod.ThemChiTieu(ctObj))
                    {
                        XtraMessageBox.Show("Thêm mới thất bại!");
                        return;
                    }
                    else
                    {
                        ctMod.themMoiGiaChiTieu(ctObj, true);
                        XtraMessageBox.Show("Thêm mới thành công!");
                    }
                }
            }
            Load_Chi_Tieu();
        }

        //Nút Nạp Lại
        private void simpleButton8_Click(object sender, EventArgs e)
        {
            them = 1;
            Load_Chi_Tieu();
        }

        //Thêm Nền Mẫu
        private void btnThemNM_Click(object sender, EventArgs e)
        {
            if (a2 == 1)
            {
                gpNM.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                a2++;
            }
            else
            {
                gpNM.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                a2 = 1;
            }
        }

        //Lưu Nền Mẫu
        private void btnLuuNM_Click_1(object sender, EventArgs e)
        {
            //gpNM.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            nmObj.Tennenmau = txtTenNM.Text;
            if (nmObj.Tennenmau == "")
            {
                XtraMessageBox.Show("Chưa điền tên nền mẫu!");
                return;
            }
            if (ctMod.KTNenMau(nmObj) <= 0)
            {
                if (ctMod.ThemNenMau(nmObj))
                {
                    XtraMessageBox.Show("Đã thêm nền mẫu");
                    btnThemNM_Click(sender, e);
                    cbNM.Properties.DataSource = ctMod.GetNenMau();
                }
            }
            else { XtraMessageBox.Show("Nền mẫu đã tồn tại"); }
            
        }

        //Thêm PP
        private void btnThemPP_Click_1(object sender, EventArgs e)
        {
            if (a3 == 1)
            {
                gpPP.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                a3++;
            }
            else
            {
                gpPP.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                a3 = 1;
            }
        }

        //Lưu PP
        private void btnLuuPP_Click_1(object sender, EventArgs e)
        {
           // gpPP.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            ppObj.Tenpp = txtTenPP.Text;
            if (ppObj.Tenpp == "")
            {
                XtraMessageBox.Show("Chưa điền tên phương pháp!");
                return;
            }
            if (ctMod.KTPhuongPhap(ppObj) <= 0)
            {
                if (ctMod.ThemPhuongPhap(ppObj))
                {
                    XtraMessageBox.Show("Đã thêm phương pháp");
                    btnThemPP_Click_1(sender, e);
                    cbPP.Properties.DataSource = ctMod.GetPhuongPhap();
                }
            }
            else { XtraMessageBox.Show("Phương pháp đã tồn tại"); }
            
        }

        //Nut Thêm Quy Chuẩn
        private void btnThemQC_Click_1(object sender, EventArgs e)
        {
            if (a4 == 1)
            {
                gpQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                a4++;
            }
            else
            {
                gpQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                a4 = 1;
            }
        }

        //Lưu Quy Chuẩn
        private void btnLuuQC_Click_1(object sender, EventArgs e)
        {
            //gpQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            qcObj.Tenqc = txtTenQC.Text;
            if (qcObj.Tenqc == "")
            {
                XtraMessageBox.Show("Chưa điền tên quy chuẩn!");
                return;
            }
            if (ctMod.KTQuyChuan(qcObj) <= 0)
            {
                if (ctMod.ThemQuyChuan(qcObj))
                {
                    XtraMessageBox.Show("Đã thêm quy chuẩn");
                    btnThemQC_Click_1(sender, e);
                    cbQC.Properties.DataSource = ctMod.GetQuyChuan();
                }
            }
            else { XtraMessageBox.Show("Quy chuẩn đã tồn tại"); }
            
        }

        //Thêm Đơn Vị
        private void btnThemDV_Click_1(object sender, EventArgs e)
        {
            if (a1 == 1)
            {
                gpDV1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                a1++;
            }
            else
            {
                gpDV1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                a1 = 1;
            }
        }

        //Lưu Đơn Vị
        private void btnLuuDV_Click_1(object sender, EventArgs e)
        {
           // gpDV1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            dvObj.Tendv = txtTenDV.Text;
            if (dvObj.Tendv=="")
            {
                XtraMessageBox.Show("Chưa điền tên đơn vị!");
                return;
            }
            if (ctMod.KTDonVi(dvObj) <= 0)
            {
                if (ctMod.ThemDonVi(dvObj))
                {
                    XtraMessageBox.Show("Đã thêm đơn vị");
                    btnThemDV_Click_1(sender, e);
                    cbDV.Properties.DataSource = ctMod.GetDonVi();

                }
            }
            else { XtraMessageBox.Show("Đơn vị đã tồn tại"); }
           
        }
        private void simpleButton32_Click_1(object sender, EventArgs e)//hủy nền mẫu
        {
            btnThemNM_Click(sender, e);
        }

        private void simpleButton40_Click(object sender, EventArgs e)//hủy phương pháp
        {
            btnThemPP_Click_1(sender, e);
        }

        private void simpleButton81_Click(object sender, EventArgs e)//hủy quy chuẩn
        {
            btnThemQC_Click_1(sender, e);
        }

        private void simpleButton82_Click(object sender, EventArgs e)//hủy đơn vị
        {
            btnThemDV_Click_1(sender, e);
        }

        #endregion QUẢN LÝ DANH MỤC CHỈ TIÊU

        #region QUẢN LÝ ĐƠN VỊ

        private void simpleButton62_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls|Excel Workbook|*.xlsm", ValidateNames = true };
            textBox3.Text = ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : "";
        }   // chọn file

        private void simpleButton46_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "Đơn Vị";
                if (ws == null)
                {
                    MessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                int fontSizeTenTruong = 14;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A1");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 5.5;

                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B1");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN ĐƠN VỊ";
                row23_MaSP.ColumnWidth = 50;

                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A1", "B1");

                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                //in đậm
                row23_CotTieuDe.Font.Bold = true;

                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);

                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A1", "B1"));

                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }   // Xuất File Mẫu

        private void simpleButton47_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" || textBox3.Text == "Chưa có tệp nào được chọn")
            {
                MessageBox.Show("Vui lòng nhập đường dẫn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox3.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] a = new string[colCount];
            int sl = 0, sl2 = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 == null)
                    {
                        a[j - 2] = "";
                    }
                    else
                    {
                        a[j - 2] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                dvObj.Tendv = a[0];

                if (!dvMod.TrungTenDV(dvObj))
                {
                    dvMod.addData(dvObj);
                    sl++;
                }
                else
                {
                    sl2++;
                }
            }
            MessageBox.Show("Đã thêm " + sl + " đơn vị vào CSDL!  --Có " + sl2 + " đơn vị đã tồn tại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.None);
            loadDonVi();
            textBox3.Text = "Chưa có tệp nào được chọn";

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }  // Hoàn Thành

        //Them DV
        private void itemDonVi_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage6;
            loadDonVi();
        }

        // nút thêm
        private void simpleButton15_Click(object sender, EventArgs e)
        {
            themsua = 4;
            enableDonVi(false);
            txtMaDV.Text = "";
            txtDonVi.Text = "";
        }

        // NÚT LƯU
        private void simpleButton14_Click(object sender, EventArgs e)
        {
            dvObj.Madv = txtMaDV.Text;
            dvObj.Tendv = txtDonVi.Text;
            if (txtDonVi.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ thông tin");
            }
            else
            {
                if (themsua == 4)
                {
                    //Trùng tên và trùng nền mẫu
                    if (dvMod.Check_TenDV(txtDonVi.Text).Rows.Count > 0)
                    {
                        XtraMessageBox.Show("Chỉ tiêu đã tồn tại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (dvMod.addData(dvObj))
                    {
                        XtraMessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadDonVi();
                        themsua = 0;
                    }
                    else
                    {
                        themsua = 0;
                        XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    if (dvMod.update(dvObj) && ppMod.update_CT(ppObj) && ppMod.update_PYC(ppObj) && ppMod.update_CN(ppObj) && ppMod.update_KQ(ppObj))
                    {
                        XtraMessageBox.Show("Cập nhật thông tin thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadDonVi();
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
            }
        }

        // TẢI LẠI
        private void simpleButton13_Click(object sender, EventArgs e)
        {
            loadDonVi();
        }


        void enableDonVi(bool e)
        {
            txtDonVi.ReadOnly = e;

            dgvDonVi.Enabled = e;
            if (e)
            {
                txtDonVi.BackColor = Color.Linen;
                layoutControlItem68.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem51.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }
            else
            {
                txtDonVi.BackColor = Color.White;
                layoutControlItem68.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem51.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            }
        }
        // Nút Sửa DV   
        private void capnhatDonVi_Click(object sender, EventArgs e)
        {
            enableDonVi(false);
        }

        private void XoaDonVi_Click(object sender, EventArgs e)
        {
            string madv = txtMaDV.Text;

            // điều kiện để xóa

            DialogResult dr = MessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (dvMod.delectData(madv))
                    XtraMessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    XtraMessageBox.Show("Xóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                return;
            loadDonVi();
        }

        private void loadDonVi()
        {
            enableDonVi(true);
            dgvDonVi.DataSource = dvMod.GetDataDV();
            txtMaDV.DataBindings.Clear();
            txtMaDV.DataBindings.Add("Text", dgvDonVi.DataSource, "MADV");
            txtDonVi.DataBindings.Clear();
            txtDonVi.DataBindings.Add("Text", dgvDonVi.DataSource, "TENDV");
        }

        #endregion QUẢN LÝ ĐƠN VỊ

        #region QUẢN LÝ DANH MỤC NƠI KIỂM

        private void itemNoiKiem_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage1;
            loadNoiKiem();
        }
        void enableNoiKiem(bool e)
        {
            dgvNoiKiem.Enabled = e;
            txtTenNoiKiem.ReadOnly = e;
            if (e)
            {
                txtTenNoiKiem.BackColor = Color.Linen;
                layoutControlItem69.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem70.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }

            else
            {
                txtTenNoiKiem.BackColor = Color.White;

                layoutControlItem69.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem70.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            }
        }
        // Nút Sửa NK

        private void capnhatNoiKiem_Click(object sender, EventArgs e)
        {
            enableNoiKiem(false);
        }

        private void xoaNoiKiem_Click(object sender, EventArgs e)
        {
            string mank = txtMaNoiKiem.Text;

            // điều kiện để xóa

            DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (nkMod.delectData(mank))
                    XtraMessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    XtraMessageBox.Show("Xóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                return;
            loadNoiKiem();
        }
        private void loadNoiKiem()
        {
            themsua = 0;
            enableNoiKiem(true);

            dgvNoiKiem.DataSource = nkMod.GetDataNK();
            txtMaNoiKiem.DataBindings.Clear();
            txtMaNoiKiem.DataBindings.Add("Text", dgvNoiKiem.DataSource, "MANK");
            txtTenNoiKiem.DataBindings.Clear();
            txtTenNoiKiem.DataBindings.Add("Text", dgvNoiKiem.DataSource, "TENNK");
        }

        //button thêm mới nơi kiểm
        private void simpleButton5_Click(object sender, EventArgs e)
        {
            themsua = 7;
            enableNoiKiem(false);
            txtMaNoiKiem.Text = "";
            txtTenNoiKiem.Text = "";
        }

        // button lưu nơi kiểm
        private void simpleButton4_Click(object sender, EventArgs e)
        {
            nkObj.Idkiem = txtMaNoiKiem.Text;
            nkObj.Tennoikiem = txtTenNoiKiem.Text;
            if (txtTenNoiKiem.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ thông tin");
            }
            else
            {
                if (themsua == 7)
                {
                    //Trùng tên và trùng nền mẫu
                    if (nkMod.Check_TenNK(txtTenNoiKiem.Text).Rows.Count > 0)
                    {
                        XtraMessageBox.Show("Chỉ tiêu đã tồn tại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (nkMod.addData(nkObj))
                    {
                        XtraMessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadNoiKiem();
                        themsua = 0;
                    }
                    else
                    {
                        themsua = 0;
                        XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    //  && nkMod.update_CT(nkObj)
                    if (nkMod.update(nkObj))
                    {
                        XtraMessageBox.Show("Cập nhật thông tin thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadNoiKiem();

                        nkMod.update_PYC(nkObj);
                        nkMod.update_CN(nkObj);
                        nkMod.update_KQ(nkObj);
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
            }
        }

        //NẠp lại nơi kiểm
        private void simpleButton3_Click(object sender, EventArgs e)
        {
            loadNoiKiem();
        }

        //button xóa nơi kiểm
        private void simpleButton32_Click(object sender, EventArgs e)
        {

        }

        #endregion QUẢN LÝ DANH MỤC NƠI KIỂM

        #region QUẢN LÝ DANH SÁCH NỀN MẪU

        private void simpleButton65_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls|Excel Workbook|*.xlsm", ValidateNames = true };

            textBox2.Text = ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : "";
        }

        private void simpleButton63_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "Nen Mau";
                if (ws == null)
                {
                    MessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                int fontSizeTenTruong = 14;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A1");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 5.5;

                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B1");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN NỀN MẪU";
                row23_MaSP.ColumnWidth = 50;

                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A1", "B1");

                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                //in đậm
                row23_CotTieuDe.Font.Bold = true;

                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);

                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A1", "B1"));

                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void simpleButton64_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "" || textBox2.Text == "Chưa có tệp nào được chọn")
            {
                MessageBox.Show("Vui lòng nhập đường dẫn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox2.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] a = new string[colCount];
            int sl = 0, sl2 = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 == null)
                    {
                        a[j - 2] = "";
                    }
                    else
                    {
                        a[j - 2] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                nmObj.Tennenmau = a[0];

                if (!nmMod.TrungTenNM(nmObj))
                {
                    nmMod.addData(nmObj);
                    sl++;
                }
                else
                {
                    sl2++;
                }
            }
            //dgvKhachHang.DataSource = dt;
            MessageBox.Show("Đã thêm " + sl + " nền mẫu vào CSDL!  --Có " + sl2 + " chỉ tiêu đã tồn tại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.None);
            loadNenMau();
            textBox2.Text = "Chưa có tệp nào được chọn";
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }

        //  Nút Sửa NM
        void enableNenMau(bool e)
        {

            dgvNenMau.Enabled = e;
            txtTenNenMau.ReadOnly = e;
            
            if (e)
            {
                txtTenNenMau.BackColor = Color.Linen;
                layoutControlItem63.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem34.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                
            }

            else
            {
                txtTenNenMau.BackColor = Color.White;
                
                layoutControlItem63.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem34.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

            }

        }

        private void capnhatNenMau_Click(object sender, EventArgs e)
        {
            enableNenMau(false);
        }

        private void simpleButton21_Click_1(object sender, EventArgs e)//thêm 
        {
            themsua = 5;
            enableNenMau(false);
            txtMaNenMau.Text = "";
            txtTenNenMau.Text = "";
        }

        private void simpleButton20_Click(object sender, EventArgs e)
        {
            nmObj.Manenmau = txtMaNenMau.Text;
            nmObj.Tennenmau = txtTenNenMau.Text;
            if (txtTenNenMau.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ thông tin");
            }
            else
            {
                if (themsua == 5)
                {
                    //Trùng tên và trùng nền mẫu
                    if (nmMod.Check_TenNM(txtTenNenMau.Text).Rows.Count > 0)
                    {
                        XtraMessageBox.Show("Chỉ tiêu đã tồn tại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (nmMod.addData(nmObj))
                    {
                        XtraMessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadNenMau();
                        themsua = 0;
                    }
                    else
                    {
                        themsua = 0;
                        XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    if (nmMod.update(nmObj) && nmMod.update_CT(nmObj) && nmMod.update_PYC(nmObj) && nmMod.update_CN(nmObj) && nmMod.update_KQ(nmObj))
                    {
                        XtraMessageBox.Show("Cập nhật thông tin thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadNenMau();
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
            }
        }

        private void simpleButton19_Click(object sender, EventArgs e)
        {
            loadNenMau();
        }

        private void btnXoaNenMau_Click(object sender, EventArgs e)//Xóa nền mẫu
        {
            string manm = txtMaNenMau.Text;

            // điều kiện để xóa

            DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (nmMod.delectData(manm))
                    XtraMessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    XtraMessageBox.Show("Xóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                return;
            loadNenMau();
        }

        private void loadNenMau()
        {
            enableNenMau(true);
            dgvNenMau.DataSource = nmMod.GetDataNM();
            txtMaNenMau.DataBindings.Clear();
            txtMaNenMau.DataBindings.Add("Text", dgvNenMau.DataSource, "MANENMAU");
            txtTenNenMau.DataBindings.Clear();
            txtTenNenMau.DataBindings.Add("Text", dgvNenMau.DataSource, "TENNENMAU");
        }

        #endregion QUẢN LÝ DANH SÁCH NỀN MẪU

        #region QUẢN LÝ DANH SÁCH QUY CHUẨN

        private void simpleButton68_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls|Excel Workbook|*.xlsm", ValidateNames = true };
            textBox1.Text = ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : "";
        }

        private void simpleButton66_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "Quy Chuẩn";
                if (ws == null)
                {
                    MessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                int fontSizeTenTruong = 14;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A1");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 5.5;

                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B1");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN QUY CHUẨN";
                row23_MaSP.ColumnWidth = 50;

                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A1", "B1");

                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                //in đậm
                row23_CotTieuDe.Font.Bold = true;

                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);

                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A1", "B1"));

                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void simpleButton67_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox1.Text == "Chưa có tệp nào được chọn")
            {
                MessageBox.Show("Vui lòng nhập đường dẫn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox1.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] a = new string[colCount];
            int sl = 0, sl2 = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 == null)
                    {
                        a[j - 2] = "";
                    }
                    else
                    {
                        a[j - 2] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                qcObj.Tenqc = a[0];

                if (!qcMod.TrungTenQC(qcObj))
                {
                    qcMod.addData(qcObj);
                    sl++;
                }
                else
                {
                    sl2++;
                }
            }
            MessageBox.Show("Đã thêm " + sl + " quy chuẩn vào CSDL!  --Có " + sl2 + " quy chuẩn đã tồn tại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.None);
            loadQuyChuan();
            textBox1.Text = "Chưa có tệp nào được chọn";
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }

        private void simpleButton27_Click(object sender, EventArgs e)
        {
            themsua = 6;
            enableQuyChuan(false);
            txtMaQuyChuan.Text = "";
            txtTenQuyChuan.Text = "";
        }

        private void simpleButton26_Click(object sender, EventArgs e)
        {
            qcObj.Maqc = txtMaQuyChuan.Text;
            qcObj.Tenqc = txtTenQuyChuan.Text;
            if (txtTenQuyChuan.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ thông tin");
            }
            else
            {
                if (themsua == 6)
                {
                    //Trùng tên và trùng nền mẫu
                    if (qcMod.Check_TenQC(txtTenQuyChuan.Text).Rows.Count > 0)
                    {
                        XtraMessageBox.Show("Quy chuẩn đã tồn tại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (qcMod.addData(qcObj))
                    {
                        XtraMessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadQuyChuan();
                        themsua = 0;
                    }
                    else
                    {
                        themsua = 0;
                        XtraMessageBox.Show("Thêm thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    if (qcMod.update(qcObj) && qcMod.update_CT(qcObj) && qcMod.update_PYC(qcObj) && qcMod.update_CN(qcObj))
                    {
                        XtraMessageBox.Show("Cập nhật thông tin thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadQuyChuan();
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
            }
        }
        void enableQuyChuan(bool e)
        {
            txtTenQuyChuan.ReadOnly = e;           
            dgvQuyChuan.Enabled = e;
            if (e)
            {
                
                txtTenQuyChuan.BackColor = Color.Linen;
                layoutControlItem272.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem21.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }

            else
            {
               
                layoutControlItem272.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem21.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                txtTenQuyChuan.BackColor = Color.White;
            }

        }
        private void loadQuyChuan()
        {
            themsua = 0;
            enableQuyChuan(true);
            dgvQuyChuan.DataSource = qcMod.GetDataQC();
            txtMaQuyChuan.DataBindings.Clear();
            txtMaQuyChuan.DataBindings.Add("Text", dgvQuyChuan.DataSource, "MAQC");
            txtTenQuyChuan.DataBindings.Clear();
            txtTenQuyChuan.DataBindings.Add("Text", dgvQuyChuan.DataSource, "TENQC");
        }

        private void simpleButton25_Click(object sender, EventArgs e)    //LOAD
        {
            loadQuyChuan();
        }



        // Nút Sửa QC

        private void capnhatQuyChuan_Click(object sender, EventArgs e)
        {
            enableQuyChuan(false);

        }

        private void xoaQuyChuan_Click(object sender, EventArgs e)
        {
            string maqc = txtMaQuyChuan.Text;

            // điều kiện để xóa

            DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (qcMod.delectData(maqc))
                    XtraMessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    XtraMessageBox.Show("Xóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                return;
            loadQuyChuan();
        }

        #endregion QUẢN LÝ DANH SÁCH QUY CHUẨN

        #region QUẢN LÝ NHẬN MẪU

        //private string manenmau = "";
        private string macd = "";
        public static string tempMahd;

        void loadHDNM()
        {
            cboHopDong.Properties.DataSource = nhanmauMod.GetHopDong();
            cboHopDong.Properties.ValueMember = "MAHD";
            cboHopDong.Properties.DisplayMember = "MAHD";
        }
        void loadChiTieuNhanMau()
        {
            cboPPNhanMau.DataSource = ppMod.GetDataPP();
            cboPPNhanMau.ValueMember = "MAPP";
            cboPPNhanMau.DisplayMember = "TENPP";

            int[] k = gridView21.GetSelectedRows();
            if (k.Length > 0)
            {
                DataRow row = gridView21.GetDataRow(k[0]);

                dgvChiTieu.DataSource = nhanmauMod.GetChiTieu2(row["tempMAMAU"].ToString());
            }
            

        }
        void tableNhanMau()
        {
            dt1 = new System.Data.DataTable();

            dt1.Columns.Add("MAPHIEU_YCPT");
            dt1.Columns.Add("tempMAMAU");
            dt1.Columns.Add("TENMAU");
            dt1.Columns.Add("MAHD");
            dt1.Columns.Add("TIME_NHANMAU_DK");
        }
        private void itemNhanMau_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage10;

            tableNhanMau();
            dgvNhanMau.DataSource = dt1;
            loadNhanMau();
            dgvChiTieu.DataSource = null;
        }

        private void rdTCT_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            RadioGroup edit = sender as RadioGroup;
            if (edit.SelectedIndex == 0)
            {
                gpDSCT1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                gpTCD1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else
            {
                gpDSCT1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                gpTCD1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

                cboChiDinh.Properties.DataSource = nhanmauMod.GetChiDinh();
                cboChiDinh.Properties.ValueMember = "MACD";
                cboChiDinh.Properties.DisplayMember = "TENCD";
            }
        }



        private void loadNhanMau()
        {
            loadHDNM();
            txtTenMau.Text = "";
            txtMoTaMau.Text = "";
            txtKhoiLuongMau.Text = "";
            txtTinhTrangMau.Text = "";
            txtThongTin.Text = "";
            txtNguonMau.Text = "";
            txtSeal.Text = "";
            txtNguoiNhan.Text = "";
            dteNgayNhanMau.EditValue = DateTime.Today;
            dateEdit3.EditValue = DateTime.Today;
        }

        private void btnXacNhan_Click(object sender, EventArgs e)
        {
            nhanmauObj.Time_nhanmaudk = Convert.ToDateTime(dteNgayNhanMau.EditValue).ToString("yyyy/MM/dd");
            nhanmauObj.Mahd = cboHopDong.Text;
            nhanmauObj.Tenmau = txtTenMau.Text;
            nhanmauObj.Motamau = txtMoTaMau.Text;
            nhanmauObj.Khoiluongmau = txtKhoiLuongMau.Text;
            nhanmauObj.Tinhtrangmau = txtTinhTrangMau.Text;
            nhanmauObj.Ttkhcungcap = txtThongTin.Text;
            nhanmauObj.Nguonmau = txtNguonMau.Text;
            nhanmauObj.Seal = txtSeal.Text;
            nhanmauObj.Nguoinhanmau = txtNguoiNhan.Text;
            nhanmauObj.Time_trabcpt_dk = Convert.ToDateTime(dateEdit3.EditValue).ToString("yyyy/MM/dd");
            nhanmauObj.TempMaMau = nhanmauMod.getTempMaMau();

            if (dteNgayNhanMau.EditValue == null|| txtKhachHang.Text == ""||txtTenMau.Text == ""||dateEdit3.EditValue==null)
            {
                XtraMessageBox.Show("Vui lòng điền đầy đủ các mục có dấu *");
                return;
            }
           

            nhanmauMod.addData(nhanmauObj);

            string ma = nhanmauMod.GetIDMau().Rows[0]["MA"].ToString();
            dt1.Rows.Add(ma, nhanmauObj.TempMaMau, txtTenMau.Text, cboHopDong.Text, Convert.ToDateTime(dteNgayNhanMau.EditValue).ToString("yyyy/MM/dd"));
            dgvNhanMau.DataSource = dt1;
            loadNhanMau();
            dgvChiTieu.DataSource = null;
        }

        private void btnLuuChiDinh_Click(object sender, EventArgs e)
        {
            if (cboChiDinh.Text == "")
            {
                MessageBox.Show("Vui lòng chọn chỉ định");
                return;
            }
            int[] tmau = gridView21.GetSelectedRows();
            if (tmau.Length == 0)
            {
                XtraMessageBox.Show("Vui lòng chọn mẫu hoặc nhập mẫu nếu chưa có mẫu trong danh sách mẫu!");
                return;
            }

            DataRow row2 = gridView21.GetDataRow(tmau[0]);

            DataRow row = nhanmauMod.CopyMau(row2["tempMAMAU"].ToString()).Rows[0];
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

            System.Data.DataTable dt = nhanmauMod.GetChiTieu_ChiDinh(macd, nhanmauObj);

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
            macd = "";
        }

        private void btnLuuChiTieu_Click(object sender, EventArgs e)
        {
            int[] chitieu = gridView19.GetSelectedRows();

            if (chitieu.Length == 0)
            {
                XtraMessageBox.Show("Vui lòng chọn chỉ tiêu!");
                return;
            }

            int[] tmau = gridView21.GetSelectedRows();
            if (tmau.Length == 0)
            {
                XtraMessageBox.Show("Vui lòng chọn mẫu hoặc nhập mẫu nếu chưa có mẫu trong danh sách mẫu!");
                return;
            }

            DataRow row2 = gridView21.GetDataRow(tmau[0]);

            DataRow row = nhanmauMod.CopyMau(row2["tempMAMAU"].ToString()).Rows[0];
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

            for (int i = 0; i < chitieu.Length; i++)
            {
                DataRow row1 = gridView19.GetDataRow(chitieu[i]);

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
                nhanmauObj.Tenqc = row1["TENQC"].ToString();
                if (nhanmauObj.Maqc == "")
                {
                    nhanmauObj.Maqc = "NULL";
                    nhanmauObj.Tenqc = "";
                }
                    
                

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

            XtraMessageBox.Show("Đã thêm chỉ tiêu cho mẫu " + nhanmauObj.Tenmau);
            //   dgvChiTieu.DataSource = nhanmauMod.GetChiTieu();
            loadChiTieuNhanMau();
        }

        private void cboHopDong_EditValueChanged(object sender, EventArgs e)
        {
            if (cboHopDong.Text == "")
                txtKhachHang.Text = "";
            else
                txtKhachHang.Text = nhanmauMod.GetTenKH(cboHopDong.EditValue.ToString());
        }


        private void cboChiDinh_EditValueChanged(object sender, EventArgs e)
        {
            if (cboChiDinh.Text != "" || cboChiDinh.EditValue != null)
            {
                int[] k = gridView20.GetSelectedRows();
                DataRow row = gridView20.GetDataRow(k[0]);
                macd = row[0].ToString();
            }
        }

        private void dgvNhanMau_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (gridView21.RowCount > 0)
                {
                    loadChiTieuNhanMau();

                }
                else
                {
                    XtraMessageBox.Show("Vui lòng nhập mẫu vào danh sách");
                }
            }
        }
        private void simpleButton38_Click(object sender, EventArgs e)
        {
            frmThemHopDong fm = new frmThemHopDong();
            fm.ShowDialog();
            loadHDNM();
            cboHopDong.Text = tempMahd;

        }

        #endregion QUẢN LÝ NHẬN MẪU

        #region QUẢN LÝ HỢP ĐỒNG
        // Nút sửa HD
        void enableHopDong(bool e)
        {
            gridControl1.Enabled = e;
            txtMaHD.ReadOnly = e;
            txtTenHD.ReadOnly = e;
            cboKH.ReadOnly = e;
            txtND.ReadOnly = e;
            txtNguoiChiuTN.ReadOnly = e;
            txtThoiHanHD.ReadOnly = e;
            txtTinhTrangHD.ReadOnly = e;
            if (e)
            {
                txtMaHD.BackColor = Color.Linen;
                txtTenHD.BackColor = Color.Linen;
                cboKH.BackColor = Color.Linen;
                txtND.BackColor = Color.Linen;
                txtNguoiChiuTN.BackColor = Color.Linen;
                txtThoiHanHD.BackColor = Color.Linen;
                txtTinhTrangHD.BackColor = Color.Linen;
            }
            else
            {
                txtMaHD.BackColor = Color.White;
                txtTenHD.BackColor = Color.White;
                cboKH.BackColor = Color.White;
                txtND.BackColor = Color.White;
                txtNguoiChiuTN.BackColor = Color.White;
                txtThoiHanHD.BackColor = Color.White;
                txtTinhTrangHD.BackColor = Color.White;
            }
        }
        private void CapNhatHopDong_Click(object sender, EventArgs e)
        {
            enableHopDong(false);
        }

        //XOA HOP DONG
        private void btnXoaHD_Click(object sender, EventArgs e)
        {
            hdObj.Mahd = txtMaHD.Text;

            DialogResult dr = XtraMessageBox.Show("Bạn chắc chấn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (hdMod.XoaHopDong(hdObj))
                {
                    XtraMessageBox.Show("Xóa hợp đồng thành công!");
                }
                else
                {
                    XtraMessageBox.Show("Xóa hợp đồng thất bại!");
                }
                Load_Hop_Dong();
            }
        }

        //LOAD HOP_DONG
        private void Load_Hop_Dong()
        {
            them = 1;
            enableHopDong(true);
            txtDoiMaHD.Hide();
            cboKH.Properties.DataSource = hdMod.GetKhachHang();
            cboKH.Properties.ValueMember = "MAKH";
            cboKH.Properties.DisplayMember = "TENKH";

            gridControl1.Enabled = true;
            gridControl1.DataSource = hdMod.GetHopDong();

            cboKH.DataBindings.Clear();
            cboKH.DataBindings.Add("EditValue", gridControl1.DataSource, "MAKH");
            txtMaHD.DataBindings.Clear();
            txtMaHD.DataBindings.Add("Text", gridControl1.DataSource, "MAHD");
            txtTenHD.DataBindings.Clear();
            txtTenHD.DataBindings.Add("Text", gridControl1.DataSource, "TENHD");
            txtThoiHanHD.DataBindings.Clear();
            txtThoiHanHD.DataBindings.Add("Text", gridControl1.DataSource, "THOIHAN");
            txtND.DataBindings.Clear();
            txtND.DataBindings.Add("Text", gridControl1.DataSource, "NOIDUNG");
            txtNguoiChiuTN.DataBindings.Clear();
            txtNguoiChiuTN.DataBindings.Add("Text", gridControl1.DataSource, "NGUOICHIUTN");
            txtTinhTrangHD.DataBindings.Clear();
            txtTinhTrangHD.DataBindings.Add("Text", gridControl1.DataSource, "TINHTRANG");
            txtDoiMaHD.DataBindings.Clear();
            txtDoiMaHD.DataBindings.Add("Text", gridControl1.DataSource, "ID");
        }

        //button Thêm HD
        private void simpleButton36_Click(object sender, EventArgs e)
        {
            them = 2;
            enableHopDong(false);

            txtMaHD.Text = string.Empty;  //ma hd
            txtNguoiChiuTN.Text = string.Empty;  // trach nhiem
            txtTenHD.Text = string.Empty; //ten hd
            txtThoiHanHD.Text = string.Empty;  //thoi han
            txtTinhTrangHD.Text = string.Empty; //tinh trang
            txtND.Text = string.Empty;  //noi dung
            cboKH.Text = string.Empty;  // khach hang
        }

        //NUT LUU
        private void simpleButton35_Click(object sender, EventArgs e)
        {
            if (txtMaHD.Text == "" || txtTenHD.Text == "" || cboKH.Text == "")
            {
                XtraMessageBox.Show("Vui lòng nhập đầy đủ các thông tin có dấu *");
                return;
            }
            hdObj.Id = txtDoiMaHD.Text;
            hdObj.Mahd = txtMaHD.Text;
            hdObj.Tenhd = txtTenHD.Text;
            hdObj.Makh = cboKH.EditValue.ToString();
            hdObj.Tenkh = cboKH.Text;
            hdObj.Thoihan = txtThoiHanHD.Text;
            hdObj.Nguoichiutn = txtNguoiChiuTN.Text;
            hdObj.Noidung = txtND.Text;
            hdObj.Tinhtrang = txtTinhTrangHD.Text;


            if (them == 1)
            {
                try
                {
                    if (!hdMod.UpdateHopDong(hdObj))
                    {
                        XtraMessageBox.Show("Cập nhật thất bại!");
                    }
                    else
                    {
                        XtraMessageBox.Show("Cập nhật thành công!");
                        Load_Hop_Dong();
                    }
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show("" + ex);
                }
            }
            else
            {
                if (them == 2)
                {
                    if (hdMod.KTHopDong(hdObj) <= 0)
                    {
                        if (hdMod.ThemHopDong(hdObj))
                        {
                            XtraMessageBox.Show("Thêm Khách Hàng thành công");
                            Load_Hop_Dong();
                        }
                        else
                        {
                            XtraMessageBox.Show("Thất bại!");
                        }
                    }
                    else
                    {
                        XtraMessageBox.Show("Hợp đồng đã tồn tại!");
                    }
                }
            }
        }

        // Nạp lại
        private void simpleButton34_Click(object sender, EventArgs e)
        {
            them = 1;
            Load_Hop_Dong();
        }

        #endregion QUẢN LÝ HỢP ĐỒNG

        #region XUẤT PHIẾU KẾT QUẢ

            #region Phiếu chưa xuất
        private void tableXuatKQ()
        {
            dtXuatKQ = new System.Data.DataTable();
            dtXuatKQ.Columns.Add("JOBNO");
            dtXuatKQ.Columns.Add("TENKH");
            dtXuatKQ.Columns.Add("DIACHI");
            dtXuatKQ.Columns.Add("TTKHCUNGCAP");
            dtXuatKQ.Columns.Add("MOTAMAU");
            dtXuatKQ.Columns.Add("SEAL");
            dtXuatKQ.Columns.Add("NGUONMAU");
            dtXuatKQ.Columns.Add("TIME_NHANMAU_TT");
            dtXuatKQ.Columns.Add("TIME_TRABCPT_DK");
            dtXuatKQ.Columns.Add("TIME_PHANTICH_DK");
            dtXuatKQ.Columns.Add("TIME_CHUYENMAU");
            dtXuatKQ.Columns.Add("MAMAU");
            dtXuatKQ.Columns.Add("TENMAU");
            dtXuatKQ.Columns.Add("MACT");
            dtXuatKQ.Columns.Add("TENCT");
            dtXuatKQ.Columns.Add("MAPP");
            dtXuatKQ.Columns.Add("TENPP");
            dtXuatKQ.Columns.Add("MADV");
            dtXuatKQ.Columns.Add("TENDV");
            dtXuatKQ.Columns.Add("LOD");
            dtXuatKQ.Columns.Add("KETQUA_PT");
            dtXuatKQ.Columns.Add("MAQC");
            dtXuatKQ.Columns.Add("TENQC");
            dtXuatKQ.Columns.Add("NHOMCT");
            dtXuatKQ.Columns.Add("KETLUAN");

            dtXuatKQ_PhuLuc = new System.Data.DataTable();
          
           
            dtXuatKQ_PhuLuc.Columns.Add("MAMAUPL");
            dtXuatKQ_PhuLuc.Columns.Add("TENMAUPL");
            dtXuatKQ_PhuLuc.Columns.Add("MACTPL");
            dtXuatKQ_PhuLuc.Columns.Add("TENCTPL");
            dtXuatKQ_PhuLuc.Columns.Add("MAPPPL");
            dtXuatKQ_PhuLuc.Columns.Add("TENPPPL");
            dtXuatKQ_PhuLuc.Columns.Add("MADVPL");
            dtXuatKQ_PhuLuc.Columns.Add("TENDVPL");
            dtXuatKQ_PhuLuc.Columns.Add("LODPL");
            dtXuatKQ_PhuLuc.Columns.Add("KETQUA_PTPL");
            dtXuatKQ_PhuLuc.Columns.Add("MAQCPL");
            dtXuatKQ_PhuLuc.Columns.Add("TENQCPL");
            dtXuatKQ_PhuLuc.Columns.Add("NHOMCTPL");
            dtXuatKQ_PhuLuc.Columns.Add("KETLUANPL");

        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {//
            if (checkEdit1.Checked == true)
                layoutKL.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            else
                layoutKL.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void checkEdit6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit6.Checked == true)
                layoutKL1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            else
                layoutKL1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void loadXuatKetQua()
        {
            gridControl19.DataSource = xkq.getJobNOChuaXuat();
            if (gridView35.RowCount > 0)
            {
                layoutControlItem168.DataBindings.Clear();
                layoutControlItem168.DataBindings.Add("Text", gridControl19.DataSource, "JOBNO");
            }
            else
                layoutControlItem168.Text = "";

            gridControl27.DataSource = xkq.getJobNODaXuat();
            layoutControlItem170.DataBindings.Clear();
            layoutControlItem170.DataBindings.Add("Text", gridControl27.DataSource, "JOBNO");
            gridControl26.DataSource = xkq.getDSmauRAW();
            if (gridView65.RowCount > 0)
            {
                layoutControlItem298.DataBindings.Clear();
                layoutControlItem298.DataBindings.Add("Text", gridControl26.DataSource, "MAMAU");
            }
            else
                layoutControlItem298.Text = "";
           
        }
        private void simpleButton61_Click(object sender, EventArgs e) //Button mẫu đã xuất
        {
            int[] k = gridView41.GetSelectedRows();
            if (k.Length > 0 && gridView43.RowCount > 0)
            {
                bool checktemp = false;
                string mau = "";
                for (int i = 0; i < k.Length; i++)
                {
                    DataRow row = gridView41.GetDataRow(k[i]);
                    if (checktemp == false)
                    {
                        mau = row[0].ToString();
                        checktemp = true;
                    }
                    else
                        mau = mau + ", " + row[0];
                }
                if (XtraMessageBox.Show("Xuất các mẫu: " + mau + "\n\nKết luận: " + textEdit36.Text, "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    tableXuatKQ();
                    for (int i = 0; i < k.Length; i++)
                    {
                        DataRow row = gridView41.GetDataRow(k[i]);

                        System.Data.DataTable dtXuat = xkq.getThongTinMauDaXuat(row["MAMAU"].ToString(), row["LANKIEMTHU"].ToString());
                        foreach (DataRow row1 in dtXuat.Rows)
                        {

                            dtXuatKQ.Rows.Add(row1["JOBNO"], row1["TENKH"], row1["DIACHI"], row1["TTKHCUNGCAP"], row1["MOTAMAU"], row1["SEAL"], row1["NGUONMAU"], row1["TIME_NHANMAU_TT"], row1["TIME_TRABCPT_DK"], row1["TIME_PHANTICH_DK"], row1["TIME_CHUYENMAU"], row1["MAMAU"], row1["TENMAU"], row1["MACT"], row1["TENCT"], row1["MAPP"], row1["TENPP"], row1["MADV"], row1["TENDV"], row1["LOD"], row1["KETQUA_PT"], row1["MAQC"], row1["TENQC"], row1["NHOMCHITIEU"], textEdit36.Text);
                            if (row1["KETQUA_PT"].ToString() != row1["KETQUAN_PTPHULUC"].ToString())
                            {
                                dtXuatKQ_PhuLuc.Rows.Add(row1["MAMAU"], row1["TENMAU"], row1["MACT"], row1["TENCTPHULUC"], row1["MAPP"], row1["TENPP"], row1["MADV"], row1["TENDV"], row1["LOD"], row1["KETQUAN_PTPHULUC"], row1["MAQC"], row1["TENQC"], row1["NHOMCHITIEU"], textEdit36.Text);
                            }

                        }
                    }
                    DataSet ds1 = new DataSet();
                    ds1.Tables.Add(dtXuatKQ);
                    ds1.Tables.Add(dtXuatKQ_PhuLuc);
                    Report.frmrpXuatKQ.dt2 = dtXuatKQ.Copy();
                    Report.rpXuatKQ.dtPhuLuc = dtXuatKQ_PhuLuc.Copy();
                    Report.frmrpXuatKQ__QC.dt2 = dtXuatKQ.Copy();
                    Report.rpXuatKQ_QC.dtPhuLuc = dtXuatKQ_PhuLuc.Copy();
                    ds1.WriteXmlSchema("ReportBaoCaoKetQua.xml");
                    if (checkEdit5.Checked == true)
                    {
                        frmrpXuatKQ__QC f = new frmrpXuatKQ__QC();
                        f.ShowDialog();
                    }
                    else
                    {
                        frmrpXuatKQ f = new frmrpXuatKQ();
                        f.ShowDialog();
                    }
                    loadXuatKetQua();
                }
                else
                    return;
            }
            else
            {
                XtraMessageBox.Show("Chưa chọn mẫu cần xuất");
            }
        }
        
        private void layoutControlItem168_TextChanged(object sender, EventArgs e)//Load mẫu chưa xuất
        {
            gridControl21.DataSource = xkq.getMau(layoutControlItem168.Text);
        }

        private void simpleButton58_Click(object sender, EventArgs e)//Button mẫu chưa xuất
        {
            int[] k = gridView37.GetSelectedRows();
            if (k.Length > 0 && gridView35.RowCount > 0)
            {
                bool checktemp = false;
                string mau = "";
                for (int i = 0; i < k.Length; i++)
                {
                    DataRow row = gridView37.GetDataRow(k[i]);
                    if (checktemp == false)
                    {
                        mau = row[0].ToString();
                        checktemp = true;
                    }
                    else
                        mau = mau+ ", " + row[0];
                }
                if (XtraMessageBox.Show("Bạn có chắc chắn xuất các mẫu: " + mau + "\n\nKết luận: " + searchLookUpEdit5.Text, "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    tableXuatKQ();
                    for (int i = 0; i < k.Length; i++)
                    {
                        DataRow row = gridView37.GetDataRow(k[i]);

                        System.Data.DataTable dtXuat = xkq.getThongTinMau(row["MAMAU"].ToString());
                        foreach (DataRow row1 in dtXuat.Rows)
                        {
                            //iNSERT CONG NO
                            cnObj.Maphieu = row1["MAPHIEU_YCPT"].ToString();
                            cnObj.Mahd = row1["MAHD"].ToString();
                            cnObj.Jobno = row1["JOBNO"].ToString();
                            cnObj.Mamau = row1["MAMAU"].ToString();
                            cnObj.Tenmau = row1["TENMAU"].ToString();
                            cnObj.Time_nhanmau_tt = Convert.ToDateTime(row1["TIME_NHANMAU_TT"].ToString()).ToString("yyyy/MM/dd");
                            cnObj.Manenmau = row1["MANENMAU"].ToString();
                            if (cnObj.Manenmau == "")
                                cnObj.Manenmau = "NULL";
                            cnObj.Nenmau = row1["TENNENMAU"].ToString();
                            cnObj.Machidinh = row1["MACD"].ToString();
                            if (cnObj.Machidinh == "")
                                cnObj.Machidinh = "NULL";
                            cnObj.Chidinh = row1["TENCD"].ToString();
                            cnObj.Machitieu = row1["MACT"].ToString();
                            cnObj.Chitieu = row1["TENCTPHULUC"].ToString();
                            cnObj.Maqc = row1["MAQC"].ToString();
                            if (cnObj.Maqc == "")
                                cnObj.Maqc = "NULL";
                            cnObj.Tenqc = row1["TENQC"].ToString();
                            cnObj.Maphuongphap = row1["MAPP"].ToString();
                            if (cnObj.Maphuongphap == "")
                                cnObj.Maphuongphap = "NULL";
                            cnObj.Phuongphap = row1["TENPP"].ToString();
                            cnObj.Dongia = row1["DONGIA"].ToString();
                            cnObj.Lankiemthu = cnMod.LanKiemThu(cnObj.Maphieu);
                            cnObj.Soluong = row1["SOLUONGKIEM"].ToString();
                            if (cnObj.Soluong == "")
                                cnObj.Soluong = "1";
                            cnObj.Lod = row1["LOD"].ToString();
                            cnObj.Madv = row1["MADV"].ToString();
                            if (cnObj.Madv == "")
                                cnObj.Madv = "NULL";

                            cnObj.Tendv = row1["TENDV"].ToString();
                            cnObj.Mank = row1["MANK"].ToString();
                            cnObj.Tennk = row1["TENNK"].ToString();
                            if (cnObj.Mank == "")
                            {
                                cnObj.Mank = "NULL";
                            }
                            cnObj.Nhomchitieu = row1["NHOMCHITIEU"].ToString();
                            cnObj.Userxuat = Login.tenUser;
                            cnObj.Makh = row1["MAKH"].ToString();
                            cnObj.Tenkh = row1["TENKH"].ToString();
                            cnObj.Ketqua = row1["KETQUA_PTPHULUC"].ToString();
                            string TENCTphuluc = row1["TENCTPHULUC"].ToString();
                            string ketquaPhuluc = row1["KETQUA_PTPHULUC"].ToString();
                            string trangthaixuat = row1["TRANGTHAI_XUAT"].ToString();

                            if (trangthaixuat != "Đã xuất")
                                cnMod.insertCongNo(cnObj);
                            //
                            dtXuatKQ.Rows.Add(row1["JOBNO"], row1["TENKH"], row1["DIACHI"], row1["TTKHCUNGCAP"], row1["MOTAMAU"], row1["SEAL"], row1["NGUONMAU"], row1["TIME_NHANMAU_TT"], row1["TIME_CHUYENMAU"], row1["MAMAU"], row1["TENMAU"], row1["MACT"], row1["TENCT"], row1["MAPP"], row1["TENPP"], row1["MADV"], row1["TENDV"], row1["LOD"], row1["KETQUA_PT"], row1["MAQC"], row1["TENQC"], row1["NHOMCHITIEU"], searchLookUpEdit5.Text);
                            if (row1["TENCT"].ToString() != row1["TENCTPHULUC"].ToString())
                            {
                                dtXuatKQ_PhuLuc.Rows.Add(row1["MAMAU"], row1["TENMAU"], row1["MACT"], row1["TENCTPHULUC"], row1["MAPP"], row1["TENPP"], row1["MADV"], row1["TENDV"], row1["LOD"], row1["KETQUA_PTPHULUC"], row1["MAQC"], row1["TENQC"], row1["NHOMCHITIEU"], searchLookUpEdit5.Text);
                            }
                        }
                        xkq.updateTrangThaiXuat(row["MAMAU"].ToString());
                    }
                    DataSet ds1 = new DataSet();
                    ds1.Tables.Add(dtXuatKQ);
                    
                    ds1.Tables.Add(dtXuatKQ_PhuLuc);
                    Report.frmrpXuatKQ.dt2 = dtXuatKQ.Copy();
                    Report.rpXuatKQ.dtPhuLuc = dtXuatKQ_PhuLuc.Copy();
                    Report.frmrpXuatKQ__QC.dt2 = dtXuatKQ.Copy();
                    Report.rpXuatKQ_QC.dtPhuLuc = dtXuatKQ_PhuLuc.Copy();
                    ds1.WriteXmlSchema("ReportBaoCaoKetQua.xml");

                    if (checkEdit2.Checked == true)
                    {
                        frmrpXuatKQ__QC f = new frmrpXuatKQ__QC();
                        f.ShowDialog();
                    }
                    else
                    {
                        frmrpXuatKQ f = new frmrpXuatKQ();
                        f.ShowDialog();
                    }
                    loadXuatKetQua();
                }
                else
                    return;
            }
            else
            {
                XtraMessageBox.Show("Chưa chọn mẫu cần xuất");
            }
        }

        private void layoutControlItem170_TextChanged(object sender, EventArgs e)//Load mẫu đã xuất
        {            
            gridControl25.DataSource = xkq.getMauDaXuat(layoutControlItem170.Text);
        }
            #endregion
       

        #endregion XUẤT PHIẾU KẾT QUẢ

        #region QUẢN LÝ CẬP NHẬT JOBNO
        private void emptySpaceItem13_TextChanged(object sender, EventArgs e)
        {
            searchLookUpEdit3.Properties.DataSource = njnMod.getNoiKiem(emptySpaceItem13.Text);
            searchLookUpEdit3.Properties.ValueMember = "MANK";
            searchLookUpEdit3.Properties.DisplayMember = "TENNK";
        }
        // LOAD FORM
        public void Load_Form_Nhap_JobNo()
        {
            layoutControlItem152.Text = "MAHD";
            gridControl7.DataSource = njnMod.GetHopDong();
            layoutControlItem152.DataBindings.Clear();
            layoutControlItem152.DataBindings.Add("Text", gridControl7.DataSource, "MAHD");



            textBox6.Hide();
            gridControl8.DataSource = njnMod.GetMauCoJobNo();
            gridControl9.DataSource = njnMod.GetMauChuaJobNo(layoutControlItem152.Text);
            emptySpaceItem13.DataBindings.Clear();
            emptySpaceItem13.DataBindings.Add("Text", gridControl8.DataSource, "JOBNO");
            //if (gridView29.RowCount == 0)
            //  textBox6.Text = "";

        }

        //NHAP JOBNO
        private void simpleButton55_Click_1(object sender, EventArgs e)
        {
            string jobno = textEdit18.Text;
            int[] k = gridView28.GetSelectedRows();    // Mảng các dòng đã check
            if (k.Length > 0)
            {
                pycptObj.Mahd = layoutControlItem152.Text;
                pycptObj.Jobno = jobno;
                if (jobno == "")
                {
                    XtraMessageBox.Show("Vui Lòng Nhập Job No");
                    return;
                }
                else
                    if (njnMod.KiemTraJobNo(pycptObj) == "0")
                    {
                        XtraMessageBox.Show("JobNo đã tồn tại ở Hợp đồng khác!");
                    }
                    else
                    {
                        if (njnMod.KiemTraJobNo(pycptObj) == "1")
                        {
                            if (XtraMessageBox.Show("JobNo đã tồn tại ở Hợp đồng, Bạn có muốn tiếp tục!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                                return;
                        }
                        try
                        {
                            for (int i = 0; i < k.Length; i++)
                            {
                                DataRow row = gridView28.GetDataRow(k[i]);
                                string tenmau = row[0].ToString();
                                string mamau = jobno + "/" + njnMod.MaMau(jobno);
                                pycptObj.Mamau = mamau;
                                pycptObj.Tenmau = tenmau;
                             //   pycptObj.Time_nhanmau_tt = DateTime.Now.ToShortDateString();
                                pycptObj.Time_nhanmau_tt = DateTime.Now.ToString("yyyy-MM-dd HH:MM");
                                pycptObj.TempMaMau = row[4].ToString();

                                if (njnMod.CapNhatJobNo(pycptObj))
                                {
                                    dtJobNoVuaNhap.Rows.Add(jobno, mamau, row[0], row[1], row[2], row[3], pycptObj.Time_nhanmau_tt, row[4]);
                                    gridControl11.DataSource = dtJobNoVuaNhap;
                                }
                            }
                            XtraMessageBox.Show("Cập nhật thành công");
                            Load_Form_Nhap_JobNo();
                            loadQRCode();
                        }
                        catch (Exception ex)
                        {
                            XtraMessageBox.Show(ex.ToString());
                        }
                    }
            }

        }

        //TextBox Mã HD

        private void layoutControlItem152_TextChanged(object sender, EventArgs e)
        {
            gridControl9.DataSource = njnMod.GetMauChuaJobNo(layoutControlItem152.Text);
            textEdit18.Text = "";
        }

        //Xóa mẫu vừa cập nhật
        private void repositoryItemButtonEdit7_Click(object sender, EventArgs e)
        {
            string message = "Bạn chắc chắn muốn xóa";
            string caption = "Lỗi cmnr";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, caption, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                int[] k = gridView29.GetSelectedRows();     // Mảng STT dòng đã check
                try
                {
                    //string mahd = layoutControlItem152.Text;
                    DataRow row = gridView29.GetDataRow(k[0]);
                    pycptObj.Jobno = row["JOBNO"].ToString();
                    //pycptObj.Tenmau = row[2].ToString();
                    // pycptObj.Mahd = mahd;
                    njnMod.Xoa_JobNo(row[7].ToString());
                    gridView29.DeleteRow(k[0]);
                    int ma = 1;
                    foreach (DataRow row1 in njnMod.GetJobNo(pycptObj.Jobno).Rows)
                    {
                        pycptObj.TempMaMau = row1[0].ToString();
                        pycptObj.Mamau = pycptObj.Jobno + '/' + ma;
                        ma++;
                        njnMod.CapNhatJobNo(pycptObj);
                        for (int i = 0; i < gridView29.RowCount; i++)
                        {
                            DataRow row2 = gridView29.GetDataRow(i);
                            if (row2["temp_MAMAU"].ToString() == pycptObj.TempMaMau)
                            {
                                gridView29.GetDataRow(i).SetField("MAMAU", pycptObj.Mamau);
                                break;
                            }
                        }

                    }
                    Load_Form_Nhap_JobNo();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.ToString());
                }
            }
        }
        //Button xoa  mẫu có JOBNO
        private void btnXoaMauDacap_Click(object sender, EventArgs e)
        {
            string message = "Bạn chắc chắn muốn xóa";
            string caption = "Lỗi cmnr";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, caption, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                int[] k = gridView27.GetSelectedRows();     // Mảng STT dòng đã check
                try
                {
                    //string mahd = layoutControlItem152.Text;
                    DataRow row = gridView27.GetDataRow(k[0]);
                    pycptObj.Jobno = row["JOBNO"].ToString();
                    //pycptObj.Tenmau = row[2].ToString();
                    //pycptObj.Mahd = mahd;
                    if (row["TRANGTHAI_CHUYEN"].ToString() != "Chưa chuyển")
                    {
                        XtraMessageBox.Show("Mẫu đã và đang xử lý, không thể xóa");
                        return;
                    }

                    njnMod.Xoa_JobNo(row["temp_MAMAU"].ToString());
                    // gridView29.DeleteRow(k[0]);
                    for (int i = 0; i < gridView29.RowCount; i++)
                    {
                        DataRow row2 = gridView29.GetDataRow(i);
                        if (row2["temp_MAMAU"].ToString() == row["temp_MAMAU"].ToString())
                        {
                            gridView29.DeleteRow(i);
                            break;
                        }

                    }
                    int ma = 1;
                    foreach (DataRow row1 in njnMod.GetJobNo(pycptObj.Jobno).Rows)
                    {
                        pycptObj.TempMaMau = row1[0].ToString();
                        pycptObj.Mamau = pycptObj.Jobno + '/' + ma;
                        ma++;
                        njnMod.CapNhatJobNo(pycptObj);

                        for (int i = 0; i < gridView29.RowCount; i++)
                        {
                            DataRow row2 = gridView29.GetDataRow(i);

                            if (row2["temp_MAMAU"].ToString() == pycptObj.TempMaMau)
                            {
                                gridView29.GetDataRow(i).SetField("MAMAU", pycptObj.Mamau);
                                break;
                            }
                        }

                    }
                    Load_Form_Nhap_JobNo();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.ToString());
                }
            }
        }

        //Xuất PYCPT
        private void simpleButton48_Click(object sender, EventArgs e)//Button xuất phiếu yêu cầu phân tích
        {
            if (checkEdit8.Checked == false)
            {
                // Xuất phiếu yêu cầu nội bộ
                int[] k = gridView27.GetSelectedRows();
                if (k.Length > 0)
                {

                    DataRow row = gridView27.GetDataRow(k[0]);
                    System.Data.DataTable dt1 = new System.Data.DataTable();
                    dt1 = njnMod.XuatPYCPT(row["JOBNO"].ToString());

                    DataSet ds = new DataSet();
                    Report.frmXuatPYCPT.dt = dt1.Copy();
                    ds.Tables.Add(dt1);
                    ds.WriteXmlSchema("XuatPYCPT.xml");
                    frmXuatPYCPT f = new frmXuatPYCPT();
                    f.Show();
                }
            }
            else
            {
                //Xuất phiếu yêu cầu phân tích NHÀ THẦU PHỤ
                if (searchLookUpEdit3.Text == "")
                {
                    XtraMessageBox.Show("Chưa chọn nơi kiểm");
                    return;
                }
                else
                {
                    int[] k = gridView27.GetSelectedRows();
                    if (k.Length > 0)
                    {
                        DataRow row1 = gridView27.GetDataRow(k[0]);

                        System.Data.DataTable dtNhaThauPhu = njnMod.getChiTieuNhaThauPhu(row1["JOBNO"].ToString(), searchLookUpEdit3.EditValue.ToString());
                        if (dtNhaThauPhu.Rows.Count > 0)
                        {
                            DataSet ds1 = new DataSet();
                            Report.frmXuatPYCPT_NTP.dt2 = dtNhaThauPhu.Copy();
                            ds1.Tables.Add(dtNhaThauPhu);
                            ds1.WriteXmlSchema("XuatPYCPT_NhaThauPhu.xml");
                            Report.frmXuatPYCPT_NTP f = new frmXuatPYCPT_NTP();
                            f.Show();
                        }
                        else
                            XtraMessageBox.Show("Không có chỉ tiêu cần kiểm ở " + searchLookUpEdit3.Text);

                    }
                }
            }
        }

        private void checkEdit8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit8.Checked == true)
                layoutControlItem205.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            else
                layoutControlItem205.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

        }

        #endregion QUẢN LÝ CẬP NHẬT JOBNO

        #region QUẢN LÝ CẬP NHẬT THÔNG TIN MẪU

        void enabledCapNhanMau(bool e)
        {
            dteNgayNhanMau_CNM.ReadOnly = e;
            txtMaHopDong_CNM.ReadOnly = e;
            txtTenMau_CNM.ReadOnly = e;
            txtMoTaMau_CNM.ReadOnly = e;
            txtNguonMau_CNM.ReadOnly = e;
            txtKhoiLuongMau_CNM.ReadOnly = e;
            txtTinhTrangMau_CNM.ReadOnly = e;
            txtSeal_CNM.ReadOnly = e;
            txtTTKHCungCap_CNM.ReadOnly = e;
            dateEdit1.ReadOnly = e;
            dgvDSMau.Enabled = e;
            txtNguoiNhan_CNM.ReadOnly = e;

            if (e == false)
            {
                layoutControlItem128.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem127.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                emptySpaceItem10.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            }
            else
            {
                layoutControlItem128.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem127.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem10.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }
        private void loadCapNhatMau()
        {
            //
            enabledCapNhanMau(true);
            //En
            dgvDSMau.DataSource = cnmMod.GetDSMau(checkEdit9.Checked);
            txtMaHopDong_CNM.Properties.DataSource = hdMod.GetHopDong();
            txtMaHopDong_CNM.Properties.ValueMember = "MAHD";
            txtMaHopDong_CNM.Properties.DisplayMember = "MAHD";

            if (gridView24.RowCount > 0)
            {
                txtMaHopDong_CNM.DataBindings.Clear();
                txtMaHopDong_CNM.DataBindings.Add("EditValue", dgvDSMau.DataSource, "MAHD");
                txtJobNo_CNM.DataBindings.Clear();
                txtJobNo_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "JOBNO");
                txtMaMau_CNM.DataBindings.Clear();
                txtMaMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "MAMAU");
                dteNgayNhanMau_CNM.DataBindings.Clear();
                dteNgayNhanMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_NHANMAU_DK");
                txtTenMau_CNM.DataBindings.Clear();
                txtTenMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TENMAU");
                txtMoTaMau_CNM.DataBindings.Clear();
                txtMoTaMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "MOTAMAU");
                txtNguonMau_CNM.DataBindings.Clear();
                txtNguonMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "NGUONMAU");
                txtKhoiLuongMau_CNM.DataBindings.Clear();
                txtKhoiLuongMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "KHOILUONGMAU");
                txtTinhTrangMau_CNM.DataBindings.Clear();
                txtTinhTrangMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TINHTRANGMAU");
                txtSeal_CNM.DataBindings.Clear();
                txtSeal_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "SEAL");
                txtTTKHCungCap_CNM.DataBindings.Clear();
                txtTTKHCungCap_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TTKHCUNGCAP");
                txtNguoiNhan_CNM.DataBindings.Clear();
                txtNguoiNhan_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "NGUOINHANMAU");
                dateEdit4.DataBindings.Clear();
                dateEdit4.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_NHANMAU_TT");
                dateEdit1.DataBindings.Clear();
                dateEdit1.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_TRABCPT_DK");
                dateEdit2.DataBindings.Clear();
                dateEdit2.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_TRABCPT_TT");
                layoutControlItem259.DataBindings.Clear();
                layoutControlItem259.DataBindings.Add("Text", dgvDSMau.DataSource, "temp_MAMAU");
                txtTrangthai.DataBindings.Clear();
                txtTrangthai.DataBindings.Add("Text", dgvDSMau.DataSource, "TRANGTHAI_CHUYEN");
                mahd_cnm = txtMaHopDong_CNM.Text;
                tenmau_cnm = txtTenMau_CNM.Text;
                ngay_cnm = Convert.ToDateTime(dteNgayNhanMau_CNM.EditValue).ToString("yyyy/MM/dd");
            }
            else
            {
                txtMaHopDong_CNM.Text = "";
                txtJobNo_CNM.Text = "";
                txtMaMau_CNM.Text = "";
                dteNgayNhanMau_CNM.EditValue = DateTime.Now;
                txtTenMau_CNM.Text = "";
                txtMoTaMau_CNM.Text = "";
                txtNguonMau_CNM.Text = "";
                txtKhoiLuongMau_CNM.Text = "";
                txtTinhTrangMau_CNM.Text = "";
                txtSeal_CNM.Text = "";
                txtTTKHCungCap_CNM.Text = "";
                txtNguoiNhan_CNM.Text = "";
             //   dateEdit4.EditValue = "";
                dateEdit4.Text = "";
                dateEdit2.Text = "";
                dateEdit1.Text = "";
                layoutControlItem259.Text = "";
            }
            mahd_cnm = txtMaHopDong_CNM.Text;
            tenmau_cnm = txtTenMau_CNM.Text;
            ngay_cnm = Convert.ToDateTime(dteNgayNhanMau_CNM.EditValue).ToString("yyyy/MM/dd");

            dgvDSChiTieu.DataSource = cnmMod.GetDSChiTieu(layoutControlItem259.Text);
        }
        private void btnLuu_CNM_Click(object sender, EventArgs e)
        {
            nhanmauObj.Mahd = txtMaHopDong_CNM.Text;
            nhanmauObj.Jobno = txtJobNo_CNM.Text;
            nhanmauObj.Mamau = txtMaMau_CNM.Text;
            nhanmauObj.Time_nhanmaudk = Convert.ToDateTime(dteNgayNhanMau_CNM.EditValue).ToString("yyyy/MM/dd");
            nhanmauObj.Tenmau = txtTenMau_CNM.Text;
            nhanmauObj.Motamau = txtMoTaMau_CNM.Text;
            nhanmauObj.Nguonmau = txtNguonMau_CNM.Text;
            nhanmauObj.Khoiluongmau = txtKhoiLuongMau_CNM.Text;
            nhanmauObj.Tinhtrangmau = txtTinhTrangMau_CNM.Text;
            nhanmauObj.Ttkhcungcap = txtTTKHCungCap_CNM.Text;
            nhanmauObj.Seal = txtSeal_CNM.Text;
            nhanmauObj.Nguoinhanmau = txtNguoiNhan_CNM.Text;
            nhanmauObj.TempMaMau = layoutControlItem259.Text;
            if (dateEdit1.EditValue == null)
                nhanmauObj.Time_trabcpt_dk = "NULL";
            else
                nhanmauObj.Time_trabcpt_dk = Convert.ToDateTime(dateEdit1.EditValue).ToString("yyyy/MM/dd");

            if (cnmMod.Update(nhanmauObj))
            {
                    loadCapNhatMau();
            }
            else
            {
                XtraMessageBox.Show("Cập nhật thất bại");
            }
        }

        private void btnLoad_CNM_Click(object sender, EventArgs e)
        {
            loadCapNhatMau();
        }

        private void dgvDSMau_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (gridView24.RowCount > 0)
                {
                    int[] k = gridView24.GetSelectedRows();
                    DataRow row = gridView24.GetDataRow(k[0]);
                    mahd_cnm = row[0].ToString();
                    tenmau_cnm = row[3].ToString();
                    ngay_cnm = row[9].ToString();
                    dgvDSChiTieu.DataSource = cnmMod.GetDSChiTieu(layoutControlItem259.Text);////
                    if (gridView24.RowCount > 0)
                    {
                        txtMaHopDong_CNM.DataBindings.Clear();
                        txtMaHopDong_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "MAHD");
                        txtJobNo_CNM.DataBindings.Clear();
                        txtJobNo_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "JOBNO");
                        txtMaMau_CNM.DataBindings.Clear();
                        txtMaMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "MAMAU");
                        dteNgayNhanMau_CNM.DataBindings.Clear();
                        dteNgayNhanMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_NHANMAU_DK");
                        txtTenMau_CNM.DataBindings.Clear();
                        txtTenMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TENMAU");
                        txtMoTaMau_CNM.DataBindings.Clear();
                        txtMoTaMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "MOTAMAU");
                        txtNguonMau_CNM.DataBindings.Clear();
                        txtNguonMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "NGUONMAU");
                        txtKhoiLuongMau_CNM.DataBindings.Clear();
                        txtKhoiLuongMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "KHOILUONGMAU");
                        txtTinhTrangMau_CNM.DataBindings.Clear();
                        txtTinhTrangMau_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TINHTRANGMAU");
                        txtSeal_CNM.DataBindings.Clear();
                        txtSeal_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "SEAL");
                        txtTTKHCungCap_CNM.DataBindings.Clear();
                        txtTTKHCungCap_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "TTKHCUNGCAP");
                        txtNguoiNhan_CNM.DataBindings.Clear();
                        txtNguoiNhan_CNM.DataBindings.Add("Text", dgvDSMau.DataSource, "NGUOINHANMAU");
                        dateEdit4.DataBindings.Clear();
                        dateEdit4.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_NHANMAU_TT");
                        dateEdit1.DataBindings.Clear();
                        dateEdit1.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_TRABCPT_DK");
                        dateEdit2.DataBindings.Clear();
                        dateEdit2.DataBindings.Add("Text", dgvDSMau.DataSource, "TIME_TRABCPT_TT");
                        layoutControlItem259.DataBindings.Clear();
                        layoutControlItem259.DataBindings.Add("Text", dgvDSMau.DataSource, "temp_MAMAU");
                        txtTrangthai.DataBindings.Clear();
                        txtTrangthai.DataBindings.Add("Text", dgvDSMau.DataSource, "TRANGTHAI_CHUYEN");
                        mahd_cnm = txtMaHopDong_CNM.Text;
                        tenmau_cnm = txtTenMau_CNM.Text;
                        ngay_cnm = Convert.ToDateTime(dteNgayNhanMau_CNM.EditValue).ToString("yyyy/MM/dd");
                    }
                }
                else
                {
                    XtraMessageBox.Show("Vui lòng nhập mẫu vào danh sách");
                }
            }
        }
        private void btnCapNhatCT_CNM_Click(object sender, EventArgs e)
        {
            if (gridView24.RowCount > 0)
            {
                frmThemChiTieu.tempMAMAU_CNM = layoutControlItem259.Text;
                frmThemChiTieu f = new frmThemChiTieu();

                f.ShowDialog();
                loadCapNhatMau();
            }
        }
        private void checkEdit9_CheckedChanged(object sender, EventArgs e)
        {
            loadCapNhatMau();
        }
        private void btnCapNhatMau_Click(object sender, EventArgs e)
        {
            if (txtTrangthai.Text != "Đã xuất KQPT")
                enabledCapNhanMau(false);
            else
                XtraMessageBox.Show("Mẫu đã xuất không thể thay đổi thông tin!");
        }

        private void btnXoaMau_Click(object sender, EventArgs e)
        {
            if (XtraMessageBox.Show("Bạn có chắc chắn xóa?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (txtJobNo_CNM.Text != "")
                {
                    XtraMessageBox.Show("Mẫu đã nhận không thể xóa, vui lòng kiểm tra lại");
                    return;
                }
                cnmMod.delete(layoutControlItem259.Text);
                loadCapNhatMau();
            }
        }

        #endregion QUẢN LÝ CẬP NHẬT THÔNG TIN MẪU

        #region NHẬP KẾT QUẢ

        private void loadDanhSachMauKQ() //Load danh sách mẫu đã nhận
        {
            layoutControlItem262.Text = "";
            gridControl12.DataSource = nkqMod.getMau();
            if (gridView30.RowCount > 0)
            {
                layoutControlItem262.DataBindings.Clear();
                layoutControlItem262.DataBindings.Add("Text", gridControl12.DataSource, "MAMAU");
            }
            
            gridControl17.DataSource = nkqMod.getChiTieuKiemLai();
            gridControl15.DataSource = nkqMod.getChiTieuMau(layoutControlItem262.Text);
            loadDVvaPP();
        }
        void loadDVvaPP()
        {
            timkiemPP.DataSource = nkqMod.getPhuongPhap();
            timkiemPP.ValueMember = "MAPP";
            timkiemPP.DisplayMember = "TENPP";
            cboDonViKQ.DataSource = dvMod.GetDataDV();
            cboDonViKQ.ValueMember = "MADV";
            cboDonViKQ.DisplayMember = "TENDV";
        }

        private void Table()//Bảng
        {
            dtNhapKQ = new System.Data.DataTable();
            dtNhapKQ.Columns.Add("MAPHIEU_YCPT");
            dtNhapKQ.Columns.Add("TENCT");
            dtNhapKQ.Columns.Add("TENNENMAU");
            dtNhapKQ.Columns.Add("MAPP");
            dtNhapKQ.Columns.Add("KETQUA_PT");
            dtNhapKQ.Columns.Add("GHICHU_PT");
            dtNhapKQ.Columns.Add("MADV");
            dtNhapKQ.Columns.Add("LOD");
        }

        private void simpleButton51_Click(object sender, EventArgs e)//Button gọi quét QR bằng camera
        {
            QRCode qr = new QRCode();
            qr.ShowDialog();
            loadDanhSachMauKQ();
        }

        private void simpleButton50_Click(object sender, EventArgs e)//Button nhận mẫu nhập thủ công
        {
            labelControl2.Text = "";
            if (textEdit17.Text != "")
            {
                System.Data.DataTable dtMamau = new System.Data.DataTable();
                dtMamau = qrMod.KiemTramau(textEdit17.Text,false);
                if (dtMamau.Rows.Count > 0)
                {
                    DataRow row = dtMamau.Rows[0];
                    if (row["MACT"].ToString() == "")
                    {
                        XtraMessageBox.Show("Mẫu chưa có chỉ tiêu, Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (row["TRANGTHAI_CHUYEN"].ToString() != "Đã chuyển")
                    {

                        if (row["TRANGTHAI_CHUYEN"].ToString() == "Hoàn thành")
                        {
                            
                            XtraMessageBox.Show("Mẫu đã kiểm hoàn thành, Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        else
                        {
                            if (row["TRANGTHAI_CHUYEN"].ToString() == "Đã xuất")
                            {
                                if (qrMod.KiemTramau(textEdit17.Text, true).Rows.Count == 0)
                                {
                                    XtraMessageBox.Show("Mẫu đã xuất kết quả, Vui lòng kiểm tra lại!");
                                    return;
                                }
                                else
                                {
                                    qrMod.updateTrangThaiChuyen(textEdit17.Text);
                                    XtraMessageBox.Show("Mẫu ' " + row[0] + " ' nhận thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    loadDanhSachMauKQ();
                                }

                            }
                            else
                            {
                                qrMod.updateTrangThaiChuyen(textEdit17.Text);
                                XtraMessageBox.Show("Mẫu ' " + row[0] + " ' nhận thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                loadDanhSachMauKQ();
                            }
                        }
                       
                    }
                    else
                    {
                        XtraMessageBox.Show("Mẫu ' " + row[0] + " ' đã nhận rồi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        loadDanhSachMauKQ();
                    }
                }
                else
                {
                    labelControl2.Text = "Mã mẫu không tồn tại, vui lòng kiểm tra lại";
                    return;
                }
            }
            else
            {
                labelControl2.Text = "Chưa điền mã mẫu";
                return;
            }
        }


        private void layoutControlItem262_TextChanged(object sender, EventArgs e)
        {

            if (layoutControlItem262.Text != "")
            {
                if (gridView30.RowCount > 0)
                {
                    gridControl15.DataSource = nkqMod.getChiTieuMau(layoutControlItem262.Text);
                    Table();
                    gridControl16.DataSource = dtNhapKQ;
                }

            }
            else
            {
                gridControl15.DataSource = null;
            }
           
        }
        private void simpleButton52_Click(object sender, EventArgs e) //Button cập nhật
        {
            loadDVvaPP();

            int[] k = gridView31.GetSelectedRows();
            if (k.Length > 0)
            {
                for (int i = 0; i < k.Length; i++)
                {
                    DataRow row = gridView31.GetDataRow(k[i]);
                    if (row["TRANGTHAI_KQ"].ToString() == "Hoàn thành")
                    {
                        XtraMessageBox.Show("Chỉ tiêu '" + row[1] + "' đã được duyệt, Vui lòng bỏ chọn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
                for (int i = 0; i < k.Length; i++)
                {
                    bool check = false;
                    DataRow row = gridView31.GetDataRow(k[i]);

                    for (int j = 0; j < dtNhapKQ.Rows.Count; j++)
                    {
                        DataRow row1 = dtNhapKQ.Rows[j];
                        if (row[0].ToString() == row1[0].ToString())
                        {
                            check = true;
                            break;
                        }
                    }
                    if (check == false)
                        dtNhapKQ.Rows.Add(row["MAPHIEU_YCPT"], row["TENCT"], row["TENNENMAU"], row["MAPP"], row["KETQUA_PT"], row["GHICHU_PT"], row["MADV"], row["LOD"]);
                    gridControl16.DataSource = dtNhapKQ;
                }
                gridView31.ClearSelection();
                
            }
        }

        private void GanKQ(int i)
        {
            pycptObj.Maphieu_ycpt = gridView32.GetRowCellValue(i, "MAPHIEU_YCPT").ToString();

            pycptObj.Mapp = gridView32.GetRowCellValue(i, "MAPP").ToString();
            pycptObj.Ketqua_pt = gridView32.GetRowCellValue(i, "KETQUA_PT").ToString();
            pycptObj.Ghichu_pt = gridView32.GetRowCellValue(i, "GHICHU_PT").ToString();
            pycptObj.Tenct = gridView32.GetRowCellValue(i, "TENCT").ToString();
            pycptObj.Madv = gridView32.GetRowCellValue(i, "MADV").ToString();
            pycptObj.Lod = gridView32.GetRowCellValue(i, "LOD").ToString();

            if (nkqMod.getTenPhuongPhap(pycptObj.Mapp).Rows.Count > 0)
            {
                DataRow row = nkqMod.getTenPhuongPhap(pycptObj.Mapp).Rows[0];
                pycptObj.Tenpp = row[0].ToString();
            }
            if (nkqMod.getTenDonVi(pycptObj.Madv).Rows.Count > 0)
            {
                DataRow row = nkqMod.getTenDonVi(pycptObj.Madv).Rows[0];
                pycptObj.Tendv = row[0].ToString();
            }
            pycptObj.User_tra_kq = Login.tenUser;

        }

        private void simpleButton54_Click(object sender, EventArgs e)
        {
            if (gridView32.RowCount > 0)
            {
                for (int i = 0; i < gridView32.RowCount; i++)
                {
                    GanKQ(i);
                    if (pycptObj.Ketqua_pt != "")
                    {
                        if (pycptObj.Mapp == "")
                        {
                            XtraMessageBox.Show("Chỉ tiêu '" + pycptObj.Tenct + "' chưa chọn phương pháp!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        else
                        {
                            if (pycptObj.Madv == "")
                            {
                                XtraMessageBox.Show("Chỉ tiêu '" + pycptObj.Tenct + "' chưa chọn đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            else
                            {
                                if (pycptObj.Lod == "")
                                {
                                    XtraMessageBox.Show("Chỉ tiêu '" + pycptObj.Tenct + "' chưa điền LOD!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return;
                                }
                            }
                        }
                    }
                    else
                    {
                        XtraMessageBox.Show("Chỉ tiêu '" + pycptObj.Tenct + "' chưa điền kết quả!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
                for (int i = 0; i < gridView32.RowCount; i++)
                {
                    GanKQ(i);
                    if (pycptObj.Lod == "")
                        pycptObj.Lod = "-";
                    nkqMod.updateKetQua(pycptObj);
                    nkqMod.thaydoiketqua(pycptObj);

                }
                layoutControlItem262_TextChanged(sender, e);
            }
        }

        private void repositoryItemButtonEdit1_Click(object sender, EventArgs e)//button xóa
        {
            int[] i = gridView32.GetSelectedRows();
            dtNhapKQ.Rows.RemoveAt(i[0]);
            gridControl16.DataSource = dtNhapKQ;
        }

        #endregion NHẬP KẾT QUẢ

        #region QUẢN LÝ DUYỆT KẾT QUẢ

        public void Load_Form_Duyet_KQ()
        {
            gridControl18.DataSource = dkqMod.GetKetQua();

            layoutControlItem157.DataBindings.Clear();
            layoutControlItem157.DataBindings.Add("Text", gridControl18.DataSource, "MAPHIEU_YCPT");  //MAPHIEU_YCPT
            textEdit19.DataBindings.Clear();
            textEdit19.DataBindings.Add("Text", gridControl18.DataSource, "MACT");      //MACT
            textEdit20.DataBindings.Clear();
            textEdit20.DataBindings.Add("Text", gridControl18.DataSource, "TENCT");      //TENCT
            textEdit21.DataBindings.Clear();
            textEdit21.DataBindings.Add("Text", gridControl18.DataSource, "TENNENMAU");      //TENNENMAU
            textEdit25.DataBindings.Clear();
            textEdit25.DataBindings.Add("Text", gridControl18.DataSource, "TENPP");      //TENPP
            textEdit27.DataBindings.Clear();
            textEdit27.DataBindings.Add("Text", gridControl18.DataSource, "KETQUA_PT");      //KETQUA_PT
            textEdit1.DataBindings.Clear();
            textEdit1.DataBindings.Add("Text", gridControl18.DataSource, "TENDV");
            textEdit4.DataBindings.Clear();
            textEdit4.DataBindings.Add("Text", gridControl18.DataSource, "LOD");
            layoutControlItem201.DataBindings.Clear();
            layoutControlItem201.DataBindings.Add("Text", gridControl18.DataSource, "MAMAU");
            textEdit3.DataBindings.Clear();
            textEdit3.DataBindings.Add("Text", gridControl18.DataSource, "GHICHU_PT");
            layoutControlItem294.DataBindings.Clear();
            layoutControlItem294.DataBindings.Add("Text", gridControl18.DataSource, "MAPP");
        }

        private void checkButton1_CheckedChanged(object sender, EventArgs e)//Duyet chỉ tiêu không đạt
        {
            if (gridView34.RowCount > 0)
            {
                pycptObj.Maphieu_ycpt = layoutControlItem157.Text;
                pycptObj.Ngay_duyet = DateTime.Now.ToString("yyyy-MM-dd");
                pycptObj.User_duyet = Login.tenUser;
                pycptObj.Ghichu_duyet = memoEdit2.Text;
                pycptObj.Mapp = layoutControlItem294.Text;
                pycptObj.Tenpp = textEdit25.Text;
                pycptObj.Ketqua_pt = textEdit27.Text;
                pycptObj.Ghichu_pt = textEdit3.Text;
                pycptObj.Ghichu_duyet = memoEdit2.Text;

                if (memoEdit2.Text == "")
                {
                    XtraMessageBox.Show("Hãy nhập nội dung không đạt!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (dkqMod.Duyet_Khong_Dat(pycptObj))
                {
                    dkqMod.thaydoiketquaduyet(pycptObj, false);
                    XtraMessageBox.Show("Duyệt thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load_Form_Duyet_KQ();
                    //emptySpaceItem14.Text = "";
                    loadDanhSachMauKQ();
                    if (gridView34.RowCount == 0)
                    {
                        layoutControlItem157.Text = "";
                        textEdit20.Text = "";
                        textEdit21.Text = "";
                        textEdit25.Text = "";
                        textEdit1.Text = "";
                        textEdit4.Text = "";
                        textEdit27.Text = "";
                        memoEdit2.Text = "";
                        textEdit3.Text = "";
                    }
                }
                else
                {
                    XtraMessageBox.Show("Thao tác thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void simpleButton56_Click(object sender, EventArgs e)//Duyet chỉ tiêu đạt
        {
            if (gridView34.RowCount > 0)
            {
                pycptObj.Maphieu_ycpt = layoutControlItem157.Text;
                pycptObj.Ngay_duyet = DateTime.Now.ToString("yyyy-MM-dd");
                pycptObj.User_duyet = Login.tenUser;
                pycptObj.Ghichu_duyet = memoEdit2.Text;
                pycptObj.Mamau = layoutControlItem201.Text;
                pycptObj.Mapp = layoutControlItem294.Text;
                pycptObj.Tenpp = textEdit25.Text;
                pycptObj.Ketqua_pt = textEdit27.Text;
                pycptObj.Ghichu_pt = textEdit3.Text;
                pycptObj.Ghichu_duyet = memoEdit2.Text;

                if (dkqMod.Duyet_Dat(pycptObj))
                {
                    dkqMod.thaydoiketquaduyet(pycptObj, true);
                    if (dkqMod.KiemTraHoanThanh(pycptObj.Mamau).Rows.Count == 0)
                        dkqMod.UpdateTrangThaiChuyen(pycptObj.Mamau,false);
                    XtraMessageBox.Show("Duyệt thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    Load_Form_Duyet_KQ();
                    
                    loadDanhSachMauKQ();
                    loadXuatKetQua();
                    if (gridView34.RowCount == 0)
                    {
                        layoutControlItem157.Text = "";
                        textEdit20.Text = "";
                        textEdit21.Text = "";
                        textEdit25.Text = "";
                        textEdit1.Text = "";
                        textEdit4.Text = "";
                        textEdit27.Text = "";
                        memoEdit2.Text = "";
                        textEdit3.Text = "";
                    }
                }
                else
                {
                    XtraMessageBox.Show("Thao tác thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
        private void simpleButton24_Click(object sender, EventArgs e)
        {
            if (gridView34.RowCount > 0)
            {
                if (XtraMessageBox.Show("Bạn có chắc chắn chuyển về bộ phận nhận mẫu?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
                else
                {
                    pycptObj.Mamau = layoutControlItem201.Text;
                    if (dkqMod.UpdateTrangThaiChuyen(pycptObj.Mamau, true))
                    {
                        XtraMessageBox.Show("Mẫu '" + pycptObj.Mamau + "' chuyển thành công!");
                        Load_Form_Duyet_KQ();

                        loadDanhSachMauKQ();
                        loadXuatKetQua();
                        if (gridView34.RowCount == 0)
                        {
                            layoutControlItem157.Text = "";
                            textEdit20.Text = "";
                            textEdit21.Text = "";
                            textEdit25.Text = "";
                            textEdit1.Text = "";
                            textEdit4.Text = "";
                            textEdit27.Text = "";
                            memoEdit2.Text = "";
                            textEdit3.Text = "";
                        }
                    }
                    else
                        XtraMessageBox.Show("Mẫu '" + pycptObj.Mamau + "' chuyển thất bại!");
                        
                }
            }
            
        }

        #endregion QUẢN LÝ DUYỆT KẾT QUẢ

        #region TẠO QR CODE

        private void loadQRCode()
        {
            gridControl4.DataSource = qrMod.getDSMauCoJobNo();
            tableQRCode();
        }

        private void tableQRCode()
        {
            dtTaoQR = new System.Data.DataTable();
            dtTaoQR.Columns.Add("MAHD");
            dtTaoQR.Columns.Add("JOBNO");
            dtTaoQR.Columns.Add("MAMAU");
            dtTaoQR.Columns.Add("TENMAU");
            dtTaoQR.Columns.Add("QR");
        }

        private void simpleButton37_Click(object sender, EventArgs e)
        {
            
                tableQRCode();
                //gridControl5.DataSource = dtTaoQR;
                int[] k = gridView42.GetSelectedRows();
                if (k.Length > 0)
                {
                    for (int i = 0; i < k.Length; i++)
                    {
                        bool check = false;
                        DataRow row = gridView42.GetDataRow(k[i]);

                        for (int j = 0; j < dtTaoQR.Rows.Count; j++)
                        {
                            DataRow row1 = dtTaoQR.Rows[j];
                            if (row[2].ToString() == row1[2].ToString())
                            {
                                check = true;
                                break;
                            }
                        }
                        if (check == false)
                            dtTaoQR.Rows.Add(row["MAHD"], row["JOBNO"], row["MAMAU"], row["TENMAU"], row["QR"]);
                    }
                    gridControl5.DataSource = dtTaoQR;
                }
            
            
        }

        private void repositoryItemButtonEdit2_Click(object sender, EventArgs e)//XÓA TỪ DÒNG
        {
            int[] i = gridView46.GetSelectedRows();
            dtTaoQR.Rows.RemoveAt(i[0]);
            gridControl5.DataSource = dtTaoQR;
        }

        private void simpleButton39_Click_1(object sender, EventArgs e)//button xóa tất cả
        {
            tableQRCode();
            gridControl5.DataSource = dtTaoQR;
        }

        private void simpleButton42_Click(object sender, EventArgs e)
        {
            if (gridView46.RowCount > 0)
            {
                if (dtTaoQR.Rows.Count > 0)
                {
                    DataSet ds = new DataSet();
                    Report.frmrpQRCode.dt1 = dtTaoQR.Copy();
                    System.Data.DataTable dttaoQR2 = dtTaoQR.Copy();
                    ds.Tables.Add(dtTaoQR);
                    ds.WriteXmlSchema("ReportTaoQRCode.xml");

                    frmrpQRCode f = new frmrpQRCode();
                    f.ShowDialog();
                    ds.Clear();
                    dtTaoQR = dttaoQR2.Copy();
                    gridControl5.DataSource = dtTaoQR;
                }

            }
            else
            {
                XtraMessageBox.Show("Danh sách xuất trống vui lòng thêm mẫu cần xuất!");
            }
            
        }

        #endregion TẠO QR CODE

        #region NGƯỜI DÙNG

        private void loadNguoiDung() //Load thông tin người dùng
        {
            gridControl3.DataSource = ndMod.GetData();
            textEdit12.DataBindings.Clear();
            textEdit12.DataBindings.Add("Text", gridControl3.DataSource, "HOTEN");
            searchLookUpEdit2.DataBindings.Clear();
            searchLookUpEdit2.DataBindings.Add("Text", gridControl3.DataSource, "NHOM");
            textEdit14.DataBindings.Clear();
            textEdit14.DataBindings.Add("Text", gridControl3.DataSource, "TENUSER");
            textEdit26.DataBindings.Clear();
            textEdit26.DataBindings.Add("Text", gridControl3.DataSource, "MATKHAU");
            gridControl3.Enabled = true;
            textEdit14.Enabled = false;
            labelControl3.Text = "";
            themNguoiDung = 1;
        }

        private void simpleButton49_Click(object sender, EventArgs e) //Button load lại thông tin người dùng
        {
            loadNguoiDung();
        }

        private void textEdit14_EditValueChanged(object sender, EventArgs e) //load quyền sử dụng
        {
            if (gridView47.RowCount > 0)
            {
                dataGridView1.DataSource = ndMod.GetQuyenSD(textEdit14.Text);
            }
            else
            {
                dataGridView1.DataSource = ndMod.GetQuyenSD(null);
            }
        }

        private void checkEdit10_CheckedChanged(object sender, EventArgs e) //Checkbox ẩn hoặc hiện mật khẩu
        {
            if (checkEdit10.Checked == true)
                textEdit26.Properties.PasswordChar = '\0';
            else
                textEdit26.Properties.PasswordChar = '*';
        }

        private void simpleButton70_Click(object sender, EventArgs e) //Button thêm mới người dùng
        {
            themNguoiDung = 2;
            textEdit12.Text = "";
            textEdit14.Text = "";
            textEdit26.Text = "";
            textEdit14.Enabled = true;
            gridControl3.Enabled = false;
        }

        private void simpleButton69_Click(object sender, EventArgs e) //Button lưu thông tin người dùng
        {
            if (textEdit12.Text == "" || textEdit14.Text == "" || textEdit26.Text == "")
            {
                labelControl3.Text = "Vui lòng điền đầy đủ thông tin";
                return;
            }
            else
            {
                ndObj.Hoten = textEdit12.Text;
                ndObj.Nhom = searchLookUpEdit2.Text;
                ndObj.Tenuser = textEdit14.Text;
                ndObj.Matkhau = textEdit26.Text;
                if (themNguoiDung == 2)//nếu thêm
                {
                    if (ndMod.Kiemtrataikhoan(ndObj.Tenuser).Rows.Count > 0)
                    {
                        labelControl3.Text = "Tài khoản đã tồn tại vui lòng điền tên khác";
                        return;
                    }
                    else
                    {
                        for (int i = 0; i < ndObj.Tenuser.Length; i++)
                        {
                            if (!System.Text.RegularExpressions.Regex.IsMatch(ndObj.Tenuser.Substring(i, 1), "^[a-zA-Z+0-9]"))
                            {
                                labelControl3.Text = "Tên tài khoản chỉ chứa các ký tự chữ cái không dấu và số, Vui lòng kiểm tra lại";
                                return;
                            }
                        }
                        for (int i = 0; i < ndObj.Matkhau.Length; i++)
                        {
                            if (!System.Text.RegularExpressions.Regex.IsMatch(ndObj.Matkhau.Substring(i, 1), "^[a-zA-Z+0-9]"))
                            {
                                labelControl3.Text = "Mật khẩu chỉ chứa các ký tự chữ cái không dấu và số, Vui lòng kiểm tra lại";
                                return;
                            }
                        }
                        if (ndMod.insertNguoiDung(ndObj))
                        {
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                string checkedCell = dataGridView1.Rows[i].Cells[2].Value.ToString();
                                DataGridViewRow row = dataGridView1.Rows[i];
                                if (checkedCell == "1")
                                {
                                    sdObj.Idquyen = row.Cells[0].Value.ToString();
                                    sdObj.Tenquyen = row.Cells[1].Value.ToString();
                                    sdObj.Tenuser = ndObj.Tenuser;
                                    ndMod.insertQuyenSD(sdObj);
                                }
                            }
                            if (XtraMessageBox.Show("Thêm mới thành công, Bạn có tiếp tục thêm", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                simpleButton70_Click(sender, e);
                            else
                                simpleButton49_Click(sender, e);

                        }

                    }

                }//end them mới
                else//Cập nhật
                {
                    for (int i = 0; i < ndObj.Matkhau.Length; i++)
                    {
                        if (!System.Text.RegularExpressions.Regex.IsMatch(ndObj.Matkhau.Substring(i, 1), "^[a-zA-Z+0-9]"))
                        {
                            labelControl3.Text = "Mật khẩu chỉ chứa các ký tự chữ cái và số, Vui lòng kiểm tra lại";
                            return;
                        }
                    }
                    if (ndMod.updateNguoiDung(ndObj))
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            bool rowAlreadyExist = false;
                            string checkedCell = dataGridView1.Rows[i].Cells[2].Value.ToString();
                            DataGridViewRow row = dataGridView1.Rows[i];
                            sdObj.Idquyen = row.Cells[0].Value.ToString();
                            sdObj.Tenquyen = row.Cells[1].Value.ToString();
                            sdObj.Tenuser = ndObj.Tenuser;
                            if (checkedCell == "1")
                            {


                                if (ndMod.kiemtraQuyenSD(sdObj.Tenuser, sdObj.Idquyen).Rows.Count > 0)
                                    rowAlreadyExist = true;
                                if (rowAlreadyExist == false)
                                    ndMod.insertQuyenSD(sdObj);
                            }
                            else
                                ndMod.deleteQuyenSD(sdObj.Tenuser, sdObj.Idquyen);
                        }
                    }
                    simpleButton49_Click(sender, e);
                }

            }
        }

        private void repositoryItemButtonEdit3_Click(object sender, EventArgs e) //Button xóa người dùng
        {
            if (XtraMessageBox.Show("Bạn có chắc chắn xóa tài khoản " + textEdit14.Text, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ndMod.deleteNguoiDung(textEdit14.Text);
            else
                return;
            simpleButton49_Click(sender, e);
        }

        #endregion

        #region XUẤT CÔNG NỢ
        System.Data.DataTable dsct1 = new System.Data.DataTable();
        System.Data.DataTable dsmk1 = new System.Data.DataTable();
        System.Data.DataTable dt = new System.Data.DataTable();
        int count, count1;
        string thoigiankt, thoigianbd;
        public static double vat1 = 0.05;
        private void loadCongNo()
        {
            cboKhach_Hang19.Properties.DataSource = cnMod.GetKH();
            cboKhach_Hang19.Properties.ValueMember = "MAKH";
            cboKhach_Hang19.Properties.DisplayMember = "TENKH";
            cboKhach_Hang19.EditValue = null;
            cboJobNo19.Properties.DataSource = cnMod.GetJobNo();
            cboJobNo19.Properties.ValueMember = "JOBNO";
            cboJobNo19.Properties.DisplayMember = "JOBNO";
            cboJobNo19.EditValue = null;
            cboHopDong_19.Properties.DataSource = cnMod.GetHD();
            cboHopDong_19.Properties.ValueMember = "MAHD";
            cboHopDong_19.Properties.DisplayMember = "TENHD";
            cboHopDong_19.EditValue = null;

            dteNgayBD.EditValue = DateTime.Today;
            dteNgayKT.EditValue = DateTime.Today;

            dgvCongNo.DataSource = null;
        }
        private void cboKhach_Hang19_EditValueChanged_1(object sender, EventArgs e)
        {

            if (cboKhach_Hang19.Text != "")
            {
                cnObj.Makh = cboKhach_Hang19.EditValue.ToString();
                cboJobNo19.Properties.DataSource = cnMod.GetJobNo(cboKhach_Hang19.EditValue.ToString());
                cboJobNo19.Properties.ValueMember = "JOBNO";
                cboJobNo19.Properties.DisplayMember = "JOBNO";
                //cnObj.Jobno = null;

                cboHopDong_19.Properties.DataSource = cnMod.GetHD(cboKhach_Hang19.EditValue.ToString());
                cboHopDong_19.Properties.ValueMember = "MAHD";
                cboHopDong_19.Properties.DisplayMember = "TENHD";
                //cnObj.Mahd = null;
                cnObj.Makh = cboKhach_Hang19.EditValue.ToString();
            }
            else
            {
                cboKhach_Hang19.Text = "";
                cnObj.Makh = null;

                cboKhach_Hang19.Properties.DataSource = cnMod.GetKH();
                cboKhach_Hang19.Properties.ValueMember = "MAKH";
                cboKhach_Hang19.Properties.DisplayMember = "TENKH";

                cboJobNo19.Properties.DataSource = cnMod.GetJobNo();
                cboJobNo19.Properties.ValueMember = "JOBNO";
                cboJobNo19.Properties.DisplayMember = "JOBNO";

                cboHopDong_19.Properties.DataSource = cnMod.GetHD();
                cboHopDong_19.Properties.ValueMember = "MAHD";
                cboHopDong_19.Properties.DisplayMember = "TENHD";
            }

        }

        private void cboHopDong_19_EditValueChanged(object sender, EventArgs e)
        {
            if (cboHopDong_19.Text != "")
            {
                cnObj.Mahd = cboHopDong_19.EditValue.ToString();

                cboJobNo19.Properties.DataSource = cnMod.GetJobNo1(cboHopDong_19.EditValue.ToString());
                cboJobNo19.Properties.ValueMember = "JOBNO";
                cboJobNo19.Properties.DisplayMember = "JOBNO";

                cboKhach_Hang19.DataBindings.Clear();
                cboKhach_Hang19.DataBindings.Add("EditValue", cnMod.GetKH(cboHopDong_19.EditValue.ToString()), "MAKH");
            }
            else
            {
                cboHopDong_19.Text = "";
                cnObj.Mahd = null;

                cboJobNo19.Properties.DataSource = cnMod.GetJobNo();
                cboJobNo19.Properties.ValueMember = "JOBNO";
                cboJobNo19.Properties.DisplayMember = "JOBNO";

                cboHopDong_19.Properties.DataSource = cnMod.GetHD();
                cboHopDong_19.Properties.ValueMember = "MAHD";
                cboHopDong_19.Properties.DisplayMember = "TENHD";

                cboKhach_Hang19.Properties.DataSource = cnMod.GetKH();
                cboKhach_Hang19.Properties.ValueMember = "MAKH";
                cboKhach_Hang19.Properties.DisplayMember = "TENKH";
            }

        }

        private void cboJobNo19_EditValueChanged_1(object sender, EventArgs e)
        {
            if (cboJobNo19.Text != "")
            {
                cnObj.Jobno = cboJobNo19.EditValue.ToString();

                cboHopDong_19.DataBindings.Clear();
                cboHopDong_19.DataBindings.Add("EditValue", cnMod.GetHD1(cboJobNo19.EditValue.ToString()), "MAHD");

                cboKhach_Hang19.DataBindings.Clear();
                cboKhach_Hang19.DataBindings.Add("EditValue", cnMod.GetKH1(cboJobNo19.EditValue.ToString()), "MAKH");
            }
            else
            {
                cboJobNo19.Text = "";
                cnObj.Jobno = null;

                cboJobNo19.Properties.DataSource = cnMod.GetJobNo();
                cboJobNo19.Properties.ValueMember = "JOBNO";
                cboJobNo19.Properties.DisplayMember = "JOBNO";

                cboHopDong_19.Properties.DataSource = cnMod.GetHD();
                cboHopDong_19.Properties.ValueMember = "MAHD";
                cboHopDong_19.Properties.DisplayMember = "TENHD";

                cboKhach_Hang19.Properties.DataSource = cnMod.GetKH();
                cboKhach_Hang19.Properties.ValueMember = "MAKH";
                cboKhach_Hang19.Properties.DisplayMember = "TENKH";


            }


        }



        private void xuatphieucongno()
        {
            Microsoft.Office.Interop.Excel.Application xcelAll = new Microsoft.Office.Interop.Excel.Application();
            xcelAll.Application.Workbooks.Add(Type.Missing);

            xcelAll.Cells[1, 1] = "CÔNG TY TNHH CÔNG NGHỆ NHONHO";
            xcelAll.Cells[1, 1].Font.Bold = true;
            xcelAll.Cells[1, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[1, 1].ColumnWidth = 4;
            xcelAll.Cells[1, 2].ColumnWidth = 15;
            xcelAll.Cells[1, 3].ColumnWidth = 12;
            xcelAll.Cells[1, 4].ColumnWidth = 30;

            xcelAll.Cells[3, 1] = "Số: 03.06.17";
            xcelAll.Cells[3, 1].Font.Bold = true;
            xcelAll.Cells[3, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[4, 1] = "PHIẾU CÔNG NỢ";
            xcelAll.Cells[4, 1].Font.Bold = true;
            xcelAll.Cells[4, 1].Font.Size = 16;
            xcelAll.Cells[4, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[5, 1] = "Tên khách hàng:";
            xcelAll.Cells[5, 1].Font.Bold = true;
            xcelAll.Cells[5, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 1] = "Địa chỉ:";
            xcelAll.Cells[6, 1].Font.Bold = true;
            xcelAll.Cells[6, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[7, 1] = "Thời gian:";
            xcelAll.Cells[7, 1].Font.Bold = true;
            xcelAll.Cells[7, 1].Font.Name = "Times New Roman";

            string tenkh = cboKhach_Hang19.Text;
            xcelAll.Cells[5, 4] = tenkh;
            xcelAll.Cells[5, 4].Font.Bold = true;
            xcelAll.Cells[5, 4].Font.Name = "Times New Roman";

            string dckh = cnMod.GetDCKH(cboKhach_Hang19.Text).Rows[0]["DIACHI"].ToString();
            xcelAll.Cells[6, 4] = dckh;
            xcelAll.Cells[6, 4].Font.Bold = true;
            xcelAll.Cells[6, 4].Font.Name = "Times New Roman";

            xcelAll.Cells[7, 4] = "Từ ngày " + dteNgayBD.Text + " đến ngày " + dteNgayKT.Text;
            xcelAll.Cells[7, 4].Font.Bold = true;
            xcelAll.Cells[7, 4].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 1] = "STT";
            xcelAll.Cells[8, 1].Font.Bold = true;
            xcelAll.Cells[8, 1].RowHeight = 100;
            xcelAll.Cells[8, 1].VerticalAlignment = 2;
            xcelAll.Cells[8, 1].HorizontalAlignment = 3;
            xcelAll.Cells[8, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 2] = "Ngày nhận mẫu";
            xcelAll.Cells[8, 2].Font.Bold = true;
            xcelAll.Cells[8, 2].VerticalAlignment = 2;
            xcelAll.Cells[8, 2].HorizontalAlignment = 3;
            xcelAll.Cells[8, 2].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 3] = "Mã số mẫu";
            xcelAll.Cells[8, 3].Font.Bold = true;
            xcelAll.Cells[8, 3].VerticalAlignment = 2;
            xcelAll.Cells[8, 3].HorizontalAlignment = 3;
            xcelAll.Cells[8, 3].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 4] = "Mô tả mẫu";
            xcelAll.Cells[8, 4].Font.Bold = true;
            xcelAll.Cells[8, 4].VerticalAlignment = 2;
            xcelAll.Cells[8, 4].HorizontalAlignment = 3;
            xcelAll.Cells[8, 4].Font.Name = "Times New Roman";

            xcelAll.get_Range("a8", "a9").Merge(false);
            xcelAll.get_Range("b8", "b9").Merge(false);
            xcelAll.get_Range("c8", "c9").Merge(false);
            xcelAll.get_Range("d8", "d9").Merge(false);

            for (int i = 0; i < dsmk1.Rows.Count; i++)
            {
                string mamau = dsmk1.Rows[i]["MAMAU"].ToString();
                string tenmau = dsmk1.Rows[i]["TENMAU"].ToString();
                cnObj.Tenmau = tenmau;
                cnObj.Mamau = mamau;
                xcelAll.Cells[i + 10, 1] = i + 1;

                xcelAll.Cells[i + 10, 2] = DateTime.Today;
                xcelAll.Cells[i + 10, 2].Font.Name = "Times New Roman";

                xcelAll.Cells[i + 10, 3] = mamau;
                xcelAll.Cells[i + 10, 3].Font.Name = "Times New Roman";

                xcelAll.Cells[i + 10, 4] = tenmau;
                xcelAll.Cells[i + 10, 4].Font.Name = "Times New Roman";

                for (int j = 0; j <= dsct1.Rows.Count; j++)
                {
                    if (j == count1)
                    {
                        Microsoft.Office.Interop.Excel.Range t1 = xcelAll.Cells[8, j + 5];
                        Microsoft.Office.Interop.Excel.Range t2 = xcelAll.Cells[9, j + 5];
                        Microsoft.Office.Interop.Excel.Range t3 = xcelAll.Cells[4, j + 5];

                        string thanhtien = cnMod.GetTT(cnObj, thoigianbd, thoigiankt).Rows[0]["TONGTIEN"].ToString();
                        xcelAll.Cells[8, j + 5] = "Thành tiền";
                        xcelAll.Cells[8, j + 5].VerticalAlignment = 2;
                        xcelAll.Cells[8, j + 5].HorizontalAlignment = 3;
                        xcelAll.Cells[8, j + 5].ColumnWidth = 15;
                        xcelAll.Cells[8, j + 5].Font.Name = "Times New Roman";

                        xcelAll.Cells[i + 10, j + 5] = thanhtien;
                        xcelAll.Cells[8, j + 5].Font.Bold = true;
                        xcelAll.get_Range(t1, t2).Merge(false);
                        xcelAll.Cells[i + 10, j + 5].Font.Name = "Times New Roman";

                        xcelAll.get_Range("a4", t3).Merge(false);
                        xcelAll.Cells[4, j + 5].HorizontalAlignment = 3;
                        xcelAll.Cells[4, j + 5].VerticalAlignment = 3;
                    }
                    else
                    {
                        string tenct = dsct1.Rows[j]["TENCT"].ToString();
                        cnObj.Machitieu = dsct1.Rows[j]["MACT"].ToString();

                        string dongia = dsct1.Rows[j]["DONGIA"].ToString();

                        xcelAll.Cells[8, j + 5].ShrinkToFit = true;
                        xcelAll.Cells[8, j + 5] = tenct;
                        xcelAll.Cells[8, j + 5].Orientation = 90;
                        xcelAll.Cells[8, j + 5].Font.Bold = true;
                        xcelAll.Cells[8, j + 5].Font.Name = "Times New Roman";

                        xcelAll.Cells[9, j + 5].ShrinkToFit = true;
                        xcelAll.Cells[9, j + 5] = dongia;
                        xcelAll.Cells[9, j + 5].Orientation = 90;
                        xcelAll.Cells[9, j + 5].Font.Name = "Times New Roman";

                        cnObj.Mamau = mamau;
                        cnObj.Dongia = dongia;
                        System.Data.DataTable sl = cnMod.GetSL(cnObj, thoigianbd, thoigiankt);
                        int countSL = sl.Rows.Count;
                        string soluong;
                        if (countSL == 0)
                        {
                            soluong = "";
                        }
                        else
                        {
                            soluong = sl.Rows[0]["SOLUONG"].ToString();
                            xcelAll.Cells[i + 10, j + 5] = soluong;
                            xcelAll.Cells[i + 10, j + 5].Font.Name = "Times New Roman";
                        }
                    }
                }
            }
            Microsoft.Office.Interop.Excel.Range c1 = xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5];
            Microsoft.Office.Interop.Excel.Range tRange = xcelAll.get_Range("a8", c1);
            tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            Microsoft.Office.Interop.Excel.Range c2 = xcelAll.Cells[dsmk1.Rows.Count + 10, 1];
            Microsoft.Office.Interop.Excel.Range c3 = xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 4];
            xcelAll.get_Range(c2, c3).Merge(false);
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1] = "Tổng phí phân tích";
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].HorizontalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].VerticalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].Font.Name = "Times New Roman";

            string tongtien = cnMod.GetTongTien(cnObj, thoigianbd, thoigiankt).Rows[0]["TONGTIEN"].ToString();

            xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 5] = tongtien;
            xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 5].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 5].Font.Name = "Times New Roman";

            Microsoft.Office.Interop.Excel.Range c4 = xcelAll.Cells[dsmk1.Rows.Count + 11, 1];
            Microsoft.Office.Interop.Excel.Range c5 = xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 4];
            xcelAll.get_Range(c4, c5).Merge(false);
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1] = "VAT " + (vat1 * 100) + "%";
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].Font.Name = "Times New Roman";
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].HorizontalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].VerticalAlignment = 3;
            // vat1 = double.Parse(cboVAT.SelectedValue.ToString());
            double vat = double.Parse(tongtien) * vat1;
            xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 5] = vat;
            xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 5].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 5].Font.Name = "Times New Roman";

            Microsoft.Office.Interop.Excel.Range c6 = xcelAll.Cells[dsmk1.Rows.Count + 12, 1];
            Microsoft.Office.Interop.Excel.Range c7 = xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 4];
            xcelAll.get_Range(c6, c7).Merge(false);
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1] = "Tổng tiền thanh toán";
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].Font.Name = "Times New Roman";
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].HorizontalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].VerticalAlignment = 3;

            double thanhtoan = double.Parse(tongtien) + vat;
            xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5] = thanhtoan;
            xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5].Font.Name = "Times New Roman";

            xcelAll.Cells[dsmk1.Rows.Count + 13, 1] = "Bằng chữ:";
            xcelAll.Cells[dsmk1.Rows.Count + 13, 1].Font.Name = "Times New Roman";
            xcelAll.Cells[dsmk1.Rows.Count + 15, 1] = "Lập phiếu";
            xcelAll.Cells[dsmk1.Rows.Count + 15, 1].Font.Name = "Times New Roman";
            xcelAll.Cells[dsmk1.Rows.Count + 15, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 15, count1 + 4] = "Ngày " + DateTime.Today.Day + " tháng " + DateTime.Today.Month + " năm " + DateTime.Today.Year;
            xcelAll.Cells[dsmk1.Rows.Count + 15, count1 + 4].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 15, count1 + 4].Font.Name = "Times New Roman";

            xcelAll.Visible = true;
        }

        private void thongke()
        {
            Microsoft.Office.Interop.Excel.Application xcelAll = new Microsoft.Office.Interop.Excel.Application();
            xcelAll.Application.Workbooks.Add(Type.Missing);

            xcelAll.Cells[1, 1].ColumnWidth = 4;
            xcelAll.Cells[1, 2].ColumnWidth = 45;
            xcelAll.Cells[1, 3].ColumnWidth = 17;
            xcelAll.Cells[1, 4].ColumnWidth = 14;
            xcelAll.Cells[1, 5].ColumnWidth = 35;
            xcelAll.Cells[1, 6].ColumnWidth = 30;
            xcelAll.Cells[1, 7].ColumnWidth = 30;
            xcelAll.Cells[1, 8].ColumnWidth = 25;
            xcelAll.Cells[1, 9].ColumnWidth = 30;
            xcelAll.Cells[1, 10].ColumnWidth = 20;
            xcelAll.Cells[1, 11].ColumnWidth = 20;

            xcelAll.Cells[2, 1] = "BÁO CÁO CÔNG NỢ KHÁCH HÀNG";
            xcelAll.Cells[2, 1].Font.Bold = true;
            Microsoft.Office.Interop.Excel.Range c1 = xcelAll.Cells[2, 1];
            Microsoft.Office.Interop.Excel.Range c2 = xcelAll.Cells[2, 9];
            xcelAll.get_Range(c1, c2).Merge(false);
            xcelAll.Cells[2, 1].HorizontalAlignment = 3;
            xcelAll.Cells[2, 1].VerticalAlignment = 3;
            xcelAll.Cells[2, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[3, 1] = "Thời gian thống kê từ: ";
            xcelAll.Cells[3, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[4, 1] = "Họ và Tên Khách hàng: ";
            xcelAll.Cells[4, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[5, 1] = "Địa chỉ Khách hàng: ";
            xcelAll.Cells[5, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[3, 3] = Convert.ToDateTime(dteNgayBD.EditValue).ToString("yyyy/MM/dd") + " đến ngày " + Convert.ToDateTime(dteNgayKT.EditValue).ToString("yyyy/MM/dd"); ;
            xcelAll.Cells[3, 3].Font.Name = "Times New Roman";

            if (cboKhach_Hang19.Text != "[EditValue is null]" && cboHopDong_19.Text != "")
            {
                string tenkh = cboKhach_Hang19.Text;
                xcelAll.Cells[4, 3] = tenkh;
                xcelAll.Cells[4, 3].Font.Name = "Times New Roman";

                string dckh = cnMod.GetDCKH(cboKhach_Hang19.Text).Rows[0]["DIACHI"].ToString();
                xcelAll.Cells[5, 3] = dckh;
                xcelAll.Cells[5, 3].Font.Name = "Times New Roman";
            }

            xcelAll.Cells[6, 1] = "STT";
            xcelAll.Cells[6, 1].Font.Bold = true;
            xcelAll.Cells[6, 1].HorizontalAlignment = 3;
            xcelAll.Cells[6, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 2] = "TÊN KHÁCH HÀNG";
            xcelAll.Cells[6, 2].Font.Bold = true;
            xcelAll.Cells[6, 2].HorizontalAlignment = 3;
            xcelAll.Cells[6, 2].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 3] = "MAHD";
            xcelAll.Cells[6, 3].Font.Bold = true;
            xcelAll.Cells[6, 3].HorizontalAlignment = 3;
            xcelAll.Cells[6, 3].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 4] = "JOBNO";
            xcelAll.Cells[6, 4].Font.Bold = true;
            xcelAll.Cells[6, 4].HorizontalAlignment = 3;
            xcelAll.Cells[6, 4].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 5] = "NGÀY NHẬN MẪU";
            xcelAll.Cells[6, 5].Font.Bold = true;
            xcelAll.Cells[6, 5].HorizontalAlignment = 3;
            xcelAll.Cells[6, 5].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 6] = "MÃ SỐ MẪU";
            xcelAll.Cells[6, 6].Font.Bold = true;
            xcelAll.Cells[6, 6].HorizontalAlignment = 3;
            xcelAll.Cells[6, 6].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 7] = "TÊN MẪU KIỂM";
            xcelAll.Cells[6, 7].Font.Bold = true;
            xcelAll.Cells[6, 7].HorizontalAlignment = 3;
            xcelAll.Cells[6, 7].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 8] = "TÊN CHỈ TIÊU KIỂM";
            xcelAll.Cells[6, 8].Font.Bold = true;
            xcelAll.Cells[6, 8].HorizontalAlignment = 3;
            xcelAll.Cells[6, 8].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 9] = "TIÊU CHUẨN/CHỈ ĐỊNH";
            xcelAll.Cells[6, 9].Font.Bold = true;
            xcelAll.Cells[6, 9].HorizontalAlignment = 3;
            xcelAll.Cells[6, 9].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 10] = "NƠI KIỂM";
            xcelAll.Cells[6, 10].Font.Bold = true;
            xcelAll.Cells[6, 10].HorizontalAlignment = 3;
            xcelAll.Cells[6, 10].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 11] = "THÀNH TIỀN";
            xcelAll.Cells[6, 11].Font.Bold = true;
            xcelAll.Cells[6, 11].HorizontalAlignment = 3;
            xcelAll.Cells[6, 11].Font.Name = "Times New Roman";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string jobno = dt.Rows[i]["JOBNO"].ToString();
                DateTime ngayxuatkq = DateTime.Today;
                string mamau = dt.Rows[i]["MAMAU"].ToString();
                string tenmau = dt.Rows[i]["TENMAU"].ToString();
                string chitieu = dt.Rows[i]["TENCT"].ToString();
                string thanhtien = dt.Rows[i]["TONGTIEN"].ToString();
                string chidinh = dt.Rows[i]["TENCD"].ToString();
                string noikiem = dt.Rows[i]["TENNK"].ToString();
                string mahd = dt.Rows[i]["MAHD"].ToString();
                string tenkh = dt.Rows[i]["TENKH"].ToString();


                xcelAll.Cells[7 + i, 1] = i + 1;
                xcelAll.Cells[7 + i, 1].HorizontalAlignment = 3;
                xcelAll.Cells[7 + i, 1].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 2] = tenkh;
                xcelAll.Cells[7 + i, 2].HorizontalAlignment = 3;
                xcelAll.Cells[7 + i, 2].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 3] = mahd;
                xcelAll.Cells[7 + i, 3].HorizontalAlignment = 3;
                xcelAll.Cells[7 + i, 3].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 4] = jobno;
                xcelAll.Cells[7 + i, 4].HorizontalAlignment = 3;
                xcelAll.Cells[7 + i, 4].WrapText = false;
                xcelAll.Cells[7 + i, 4].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 5] = ngayxuatkq;
                xcelAll.Cells[7 + i, 5].HorizontalAlignment = 3;
                xcelAll.Cells[7 + i, 5].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 6] = mamau;
                xcelAll.Cells[7 + i, 6].HorizontalAlignment = 3;
                xcelAll.Cells[7 + i, 6].WrapText = false;
                xcelAll.Cells[7 + i, 6].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 7] = tenmau;
                xcelAll.Cells[7 + i, 7].WrapText = false;
                xcelAll.Cells[7 + i, 7].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 8] = chitieu;
                xcelAll.Cells[7 + i, 8].WrapText = false;
                xcelAll.Cells[7 + i, 8].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 9] = chidinh;
                xcelAll.Cells[7 + i, 9].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 10] = noikiem;
                xcelAll.Cells[7 + i, 10].Font.Name = "Times New Roman";

                xcelAll.Cells[7 + i, 11] = thanhtien;
                xcelAll.Cells[7 + i, 11].Font.Name = "Times New Roman";
                xcelAll.Cells[7 + i, 11].NumberFormat = "#.##0,00 ₫";
                xcelAll.Cells[7 + i, 11].Font.Name = "Times New Roman";

            }

            xcelAll.Cells[count + 7, 1] = "Tổng tiền công nợ:";
            xcelAll.Cells[count + 7, 1].Font.Bold = true;
            xcelAll.Cells[count + 7, 1].Font.Name = "Times New Roman";
            Microsoft.Office.Interop.Excel.Range c3 = xcelAll.Cells[count + 7, 1];
            Microsoft.Office.Interop.Excel.Range c4 = xcelAll.Cells[count + 7, 10];
            xcelAll.get_Range(c3, c4).Merge(false);
            xcelAll.Cells[count + 7, 1].HorizontalAlignment = 3;
            xcelAll.Cells[count + 7, 1].VerticalAlignment = 3;

            string tongtien = cnMod.GetTongTien(cnObj, thoigianbd, thoigiankt).Rows[0]["TONGTIEN"].ToString();
            xcelAll.Cells[count + 7, 11] = tongtien;
            xcelAll.Cells[count + 7, 11].Font.Bold = true;
            xcelAll.Cells[count + 7, 11].NumberFormat = "#.##0,00 ₫";
            xcelAll.Cells[count + 7, 11].Font.Name = "Times New Roman";

            Microsoft.Office.Interop.Excel.Range c5 = xcelAll.Cells[count + 7, 11];
            Microsoft.Office.Interop.Excel.Range tRange = xcelAll.get_Range("a6", c5);
            tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            xcelAll.Visible = true;
        }



        private void mi1_Click(object sender, EventArgs e)
        {
            if (cboKhach_Hang19.Text.Equals(""))
            {
                MessageBox.Show("Chưa chọn khách hàng");
            }
            else
            {
                thoigianbd = Convert.ToDateTime(dteNgayBD.EditValue).ToString("yyyy/MM/dd");
                thoigiankt = Convert.ToDateTime(dteNgayKT.EditValue).ToString("yyyy/MM/dd");
                dsct1 = cnMod.GetCT(cnObj, thoigianbd, thoigiankt);
                count1 = dsct1.Rows.Count;
                dsmk1 = cnMod.GetTM(cnObj, thoigianbd, thoigiankt);
                if (count1 == 0)
                {
                    XtraMessageBox.Show("Không có thông tin để xuất");
                }
                else
                {

                    VAT v = new VAT();
                    v.ShowDialog();
                    xuatphieucongno();
                }
            }

        }

        private void mi2_Click(object sender, EventArgs e)
        {
            thoigianbd = Convert.ToDateTime(dteNgayBD.EditValue).ToString("yyyy/MM/dd");
            thoigiankt = Convert.ToDateTime(dteNgayKT.EditValue).ToString("yyyy/MM/dd");
            dt = cnMod.GetData(cnObj, thoigianbd, thoigiankt);
            count = dt.Rows.Count;

            if (count.Equals(0))
            {
                XtraMessageBox.Show("Không có thông tin để xuất");
            }
            else
            {
                thongke();
            }
        }

        private void dgvCongNo_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu m = new ContextMenu();
                System.Windows.Forms.MenuItem addDevice = new System.Windows.Forms.MenuItem("Xuất công nợ khách hàng");
                System.Windows.Forms.MenuItem addDevice1 = new System.Windows.Forms.MenuItem("Xuất thống kê");
                m.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { addDevice, addDevice1 });
                dgvCongNo.ContextMenu = m;
                m.Show((Control)(sender), e.Location);

                addDevice.Click += new EventHandler(mi1_Click);
                addDevice1.Click += new EventHandler(mi2_Click);
            }
            else
            {
                int[] i = gridView45.GetSelectedRows();
                if (i.Length != 0)
                {
                    DataRow row = gridView45.GetDataRow(i[0]);
                    cnObj.Dongia = row[6].ToString();
                    cnObj.Macno = row[10].ToString();
                    cnObj.Mamau = row[1].ToString();
                    cnMod.updateDonGia(cnObj);
                }
            }
        }

        private void dgvCongNo_KeyUp_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int[] i = gridView45.GetSelectedRows();
                DataRow row = gridView45.GetDataRow(i[0]);
                cnObj.Dongia = row[6].ToString();
                cnObj.Macno = row[10].ToString();
                cnObj.Mamau = row[1].ToString();

                cnMod.updateDonGia(cnObj);
                btnThongKe_Click(sender, e);

            }
        }
        private void btnThongKe_Click(object sender, EventArgs e)
        {
            thoigianbd = Convert.ToDateTime(dteNgayBD.EditValue).ToString("yyyy/MM/dd");
            thoigiankt = Convert.ToDateTime(dteNgayKT.EditValue).ToString("yyyy/MM/dd");
            dgvCongNo.DataSource = cnMod.GetData(cnObj, thoigianbd, thoigiankt);
            if (gridView45.RowCount == 0)
                XtraMessageBox.Show("Không có công nợ trong khoảng thời gian này");
        }


        #endregion 

        #region NHÂN BẢN THEO JOBNO

        private void simpleButton71_Click(object sender, EventArgs e)
        {
            tempMahd = "";
            frmThemHopDong f = new frmThemHopDong();
            f.ShowDialog();
            loadHopDongNhanBan();
            if (tempMahd != "")
                searchLookUpEdit4.EditValue = tempMahd;
        }

        private void simpleButton75_Click_1(object sender, EventArgs e)
        {

            int[] k = gridView51.GetSelectedRows();

            if (k.Length > 0)
            {
                if (searchLookUpEdit4.Text == "")
                {
                    XtraMessageBox.Show("Vui lòng chọn hợp đồng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string ktra = nbjnMod.KiemTraJobNo(textEdit11.Text, searchLookUpEdit4.Text);
                if (ktra == "0")
                {
                    XtraMessageBox.Show("JobNo đã tồn tại trong hợp đồng khác, Vui lòng điền JobNo khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    if (ktra == "1")
                    {
                        if (XtraMessageBox.Show("JobNo đã tồn tại trong hợp đồng,Bạn có muốn tiếp tục!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                            return;
                    }
                    try
                    {
                        for (int i = 0; i < k.Length; i++)
                        {
                            DataRow row = gridView51.GetDataRow(k[i]);

                            nhanmauObj.Mahd = searchLookUpEdit4.Text;

                            nhanmauObj.Jobno = textEdit11.Text;
                            if (nhanmauObj.Jobno == "")
                                nhanmauObj.Mamau = "";
                            else
                                nhanmauObj.Mamau = nhanmauObj.Jobno + "/" + njnMod.MaMau(nhanmauObj.Jobno);
                            nhanmauObj.Tenmau = row["TENMAU"].ToString();
                            nhanmauObj.Time_nhanmaudk = Convert.ToDateTime(row["TIME_NHANMAU_DK"].ToString()).ToString("yyyy/MM/dd");
                            nhanmauObj.Motamau = row["MOTAMAU"].ToString();
                            nhanmauObj.Nguonmau = row["NGUONMAU"].ToString();
                            nhanmauObj.Khoiluongmau = row["KHOILUONGMAU"].ToString();
                            nhanmauObj.Tinhtrangmau = row["TINHTRANGMAU"].ToString();
                            nhanmauObj.Nguoinhanmau = row["NGUOINHANMAU"].ToString();
                            nhanmauObj.Seal = row["SEAL"].ToString();
                            if (row["TIME_TRABCPT_DK"].ToString() != "")
                                nhanmauObj.Time_trabcpt_dk = Convert.ToDateTime(row["TIME_TRABCPT_DK"].ToString()).ToString("yyyy/MM/dd");

                            nhanmauObj.Ttkhcungcap = row["TTKHCUNGCAP"].ToString();
                            nhanmauObj.TempMaMau = nhanmauMod.getTempMaMau();

                            System.Data.DataTable dtct = new System.Data.DataTable();
                            dtct = nbjnMod.GetChiTieu(row["MAMAU"].ToString());
                            foreach (DataRow row1 in dtct.Rows)
                            {
                                nhanmauObj.Mact = row1[0].ToString();
                                nhanmauObj.Tenct = row1[1].ToString();
                                nhanmauObj.Macd = row1[2].ToString();
                                nhanmauObj.Tencd = row1[3].ToString();
                                nhanmauObj.Manenmau = row1[4].ToString();
                                nhanmauObj.Tennenmau = row1[5].ToString();
                                nhanmauObj.Madv = row1[6].ToString();
                                nhanmauObj.Tendv = row1[7].ToString();
                                nhanmauObj.Lod = row1[8].ToString();
                                nhanmauObj.Mapp = row1[9].ToString();
                                nhanmauObj.Tenpp = row1[10].ToString();
                                nhanmauObj.Dongia = row1[11].ToString();
                                nhanmauObj.Nhomct = row1[12].ToString();
                                nhanmauObj.Maqc = row1[13].ToString();
                                nhanmauObj.Tenqc = row1[14].ToString();
                                nbjnMod.addData(nhanmauObj);
                            }

                            dtNhanBan.Rows.Add(textEdit13.Text, nhanmauObj.Mahd, nhanmauObj.Jobno, nhanmauObj.Mamau, nhanmauObj.Tenmau, nhanmauObj.Time_nhanmaudk, nhanmauObj.Motamau, nhanmauObj.Nguonmau, nhanmauObj.Khoiluongmau, nhanmauObj.Tinhtrangmau, nhanmauObj.Nguoinhanmau, nhanmauObj.Seal, nhanmauObj.Time_trabcpt_dk, nhanmauObj.Ttkhcungcap, nhanmauObj.TempMaMau);
                        }
                        XtraMessageBox.Show("Nhân bản thành công");
                        loadNhanBanJOBNO();
                        gridControl30.DataSource = dtNhanBan;
                    }
                    catch (Exception ex)
                    {
                        XtraMessageBox.Show(ex.ToString());
                    }
                }

            }
            else
            {
                XtraMessageBox.Show("Chưa chọn mẫu cần nhân bản");
            }

        }

        private void simpleButton74_Click_1(object sender, EventArgs e)
        {
            loadNhanBanJOBNO();
        }

        private void searchLookUpEdit4_EditValueChanged(object sender, EventArgs e)
        {
            if (searchLookUpEdit4.Text == "")
                textEdit13.Text = "";
            else
                textEdit13.Text = nhanmauMod.GetTenKH(searchLookUpEdit4.EditValue.ToString());
        }

        private void simpleLabelItem2_TextChanged_1(object sender, EventArgs e)
        {
            searchLookUpEdit4.EditValue = simpleLabelItem2.Text;
        }

        private void simpleLabelItem1_TextChanged_1(object sender, EventArgs e)
        {
            gridControl28.DataSource = nbjnMod.GetMau(simpleLabelItem1.Text);
            textEdit11.Text = simpleLabelItem1.Text;
        }

        private void repositoryItemButtonEdit5_Click(object sender, EventArgs e)
        {

            string message = "Bạn chắc chắn muốn xóa";
            string caption = "Lỗi cmnr";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, caption, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                int[] k = gridView55.GetSelectedRows();     // Mảng STT dòng đã check
                try
                {
                    //string mahd = textBoxDSMau.Text;
                    DataRow row = gridView55.GetDataRow(k[0]);
                    pycptObj.Jobno = row["JOBNO"].ToString();
                    pycptObj.Tenmau = row["TENMAU"].ToString();
                    pycptObj.Mahd = row["MAHD"].ToString();
                    nbjnMod.delMau(row["temp_MAMAU"].ToString());
                    gridView55.DeleteRow(k[0]);
                    int ma = 1;
                    foreach (DataRow row1 in njnMod.GetJobNo(pycptObj.Jobno).Rows)
                    {
                        pycptObj.TempMaMau = row1[0].ToString();
                        pycptObj.Mamau = pycptObj.Jobno + "/" + ma;
                        ma++;
                        njnMod.CapNhatJobNo(pycptObj);
                        for (int i = 0; i < gridView55.RowCount; i++)
                        {
                            DataRow row2 = gridView55.GetDataRow(i);
                            if (row2["temp_MAMAU"].ToString() == pycptObj.TempMaMau)
                            {
                                gridView55.GetDataRow(i).SetField("MAMAU", pycptObj.Mamau);
                                break;
                            }
                        }

                    }
                    loadNhanBanJOBNO();

                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.ToString());
                }
            }

        }
        void loadHopDongNhanBan()
        {
            searchLookUpEdit4.Properties.DataSource = hdMod.GetHopDong();
            searchLookUpEdit4.Properties.ValueMember = "MAHD";
            searchLookUpEdit4.Properties.DisplayMember = "MAHD";
        }
        void loadNhanBanJOBNO()
        {

            loadHopDongNhanBan();

            gridControl14.DataSource = nbjnMod.GetJOBNO();
            simpleLabelItem2.DataBindings.Clear();
            simpleLabelItem2.DataBindings.Add("Text", gridControl14.DataSource, "MAHD");
            simpleLabelItem1.DataBindings.Clear();
            simpleLabelItem1.DataBindings.Add("Text", gridControl14.DataSource, "JOBNO");

            searchLookUpEdit4.EditValue = simpleLabelItem2.Text;
            gridControl28.DataSource = nbjnMod.GetMau(simpleLabelItem1.Text);
            textEdit11.Text = simpleLabelItem1.Text;

        }


        #endregion

        #region XUẤT PHIEUS YCPT

        void loadXuatPhieu_YCPT()
        {
            
            dgvPhiey_YCPT.DataSource = njnMod.GetMauCoJobNo();
            if (gridView44.RowCount > 0)
            {
                textEdit2.DataBindings.Clear();
                textEdit2.DataBindings.Add("Text", dgvPhiey_YCPT.DataSource, "JOBNO");
            }           
            
        }
        private void textEdit2_TextChanged(object sender, EventArgs e)
        {
            if (textEdit2.Text != "")
            {
                cboNoiKiem.Properties.DataSource = njnMod.getNoiKiem(textEdit2.Text);
                cboNoiKiem.Properties.ValueMember = "MANK";
                cboNoiKiem.Properties.DisplayMember = "TENNK";
            }
            else
                cboNoiKiem.Properties.DataSource = null;

        }


        private void btnLocNoiKiem_Click(object sender, EventArgs e)
        {
            if (textEdit2.Text == "")
            {
                XtraMessageBox.Show("Danh sách trống");
                return;
            }
            else
            {
                if (cboNoiKiem.Text == "")
                {
                    dt1 = njnMod.XuatPYCPT(textEdit2.Text);

                    DataSet ds = new DataSet();
                    Report.frmXuatPYCPT.dt = dt1.Copy();
                    ds.Tables.Add(dt1);
                    ds.WriteXmlSchema("XuatPYCPT.xml");
                    frmXuatPYCPT f = new frmXuatPYCPT();
                    f.Show();
                }
                else
                {
                    System.Data.DataTable dtNhaThauPhu = njnMod.getChiTieuNhaThauPhu(textEdit2.Text, cboNoiKiem.EditValue.ToString());
                    if (dtNhaThauPhu.Rows.Count > 0)
                    {
                        DataSet ds1 = new DataSet();
                        Report.frmXuatPYCPT_NTP.dt2 = dtNhaThauPhu.Copy();
                        ds1.Tables.Add(dtNhaThauPhu);
                        ds1.WriteXmlSchema("XuatPYCPT_NhaThauPhu.xml");
                        Report.frmXuatPYCPT_NTP f = new frmXuatPYCPT_NTP();
                        f.Show();
                    }
                    else
                        XtraMessageBox.Show("Không có chỉ tiêu cần kiểm ở " + cboNoiKiem.Text);
                }

            }

            
        }
        #endregion

        #region NHÂN BẢN THEO MẪU
        //Load nhân bản theo mẫu
        void loadNhanBanMau()
        {
            layoutControlItem258.Text = "";
            gridControl29.DataSource = nbmMod.GetMau();
            if (gridView52.RowCount > 0)
            {
                layoutControlItem258.DataBindings.Clear();
                layoutControlItem258.DataBindings.Add("Text", gridControl29.DataSource, "temp_MAMAU");
            }
            else
                layoutControlItem258.Text = "";
            loadHopDongNhanBanMau();
        }
        void loadHopDongNhanBanMau()
        {
            textEdit31.Properties.DataSource = hdMod.GetHopDong();
            textEdit31.Properties.ValueMember = "MAHD";
            textEdit31.Properties.DisplayMember = "MAHD";
        }

        private void simpleButton78_Click(object sender, EventArgs e)//Thêm mới hợp đồng
        {
            tempMahd = "";
            frmThemHopDong f = new frmThemHopDong();
            f.ShowDialog();
            loadHopDongNhanBanMau();
            if (tempMahd != "")
                textEdit31.EditValue = tempMahd;
        }

        private void textEdit31_EditValueChanged(object sender, EventArgs e)
        {
            if (textEdit31.Text == "")
                textEdit32.Text = "";
            else
                textEdit32.Text = nhanmauMod.GetTenKH(textEdit31.EditValue.ToString());
        }

        private void layoutControlItem258_TextChanged(object sender, EventArgs e)
        {
            System.Data.DataTable dttemp = new System.Data.DataTable();
            dttemp = nbmMod.LayThongTinMau(layoutControlItem258.Text);
            if (dttemp.Rows.Count > 0)
            {
                textEdit31.EditValue = dttemp.Rows[0]["MAHD"].ToString();
                textEdit33.Text = dttemp.Rows[0]["JOBNO"].ToString();
                textEdit29.Text = dttemp.Rows[0]["MAMAU"].ToString();
                textEdit30.Text = dttemp.Rows[0]["TENMAU"].ToString();
                textEdit34.Text = dttemp.Rows[0]["TIME_NHANMAU_DK"].ToString();
                textEdit40.Text = dttemp.Rows[0]["KHOILUONGMAU"].ToString();
                textEdit41.Text = dttemp.Rows[0]["NGUONMAU"].ToString();
                textEdit38.Text = dttemp.Rows[0]["MOTAMAU"].ToString();
                textEdit43.Text = dttemp.Rows[0]["TIME_TRABCPT_DK"].ToString();
                textEdit28.Text = dttemp.Rows[0]["NGUOINHANMAU"].ToString();
                textEdit39.Text = dttemp.Rows[0]["TINHTRANGMAU"].ToString();
                textEdit42.Text = dttemp.Rows[0]["TTKHCUNGCAP"].ToString();
                textEdit37.Text = dttemp.Rows[0]["SEAL"].ToString();
                gridControl13.DataSource = nbmMod.LayDSchiTieuMau(layoutControlItem258.Text);
            }

        }

        private void simpleButton77_Click(object sender, EventArgs e)
        {
            loadNhanBanMau();
        }

        private void simpleButton76_Click(object sender, EventArgs e)
        {
            int[] k = gridView58.GetSelectedRows();
            if (k.Length == 0)
            {
                XtraMessageBox.Show("Chưa chọn chỉ tiêu");
                return;
            }
            else
            {
                if (textEdit31.Text == "")
                {
                    XtraMessageBox.Show("Vui lòng chọn hợp đồng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string ktra = nbjnMod.KiemTraJobNo(textEdit33.Text, textEdit31.Text);
                if (ktra == "0")
                {
                    XtraMessageBox.Show("JobNo đã tồn tại trong hợp đồng khác, Vui lòng điền JobNo khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    if (ktra == "1")
                    {
                        if (XtraMessageBox.Show("JobNo đã tồn tại trong hợp đồng,Bạn có muốn tiếp tục!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                            return;
                    }
                    try
                    {
                        nhanmauObj = new Phieu_Nhan_Mau_Obj();
                        nhanmauObj.Mahd = textEdit31.Text;
                        nhanmauObj.Jobno = textEdit33.Text;
                        if (nhanmauObj.Jobno == "")
                            nhanmauObj.Mamau = "";
                        else
                            nhanmauObj.Mamau = nhanmauObj.Jobno + "/" + njnMod.MaMau(nhanmauObj.Jobno);
                        nhanmauObj.Tenmau = textEdit30.Text;
                        nhanmauObj.Time_nhanmaudk = Convert.ToDateTime(textEdit34.Text).ToString("yyyy/MM/dd");
                        nhanmauObj.Khoiluongmau = textEdit40.Text;
                        nhanmauObj.Nguonmau = textEdit41.Text;
                        nhanmauObj.Motamau = textEdit38.Text;
                        if (textEdit43.Text != "")
                            nhanmauObj.Time_trabcpt_dk = Convert.ToDateTime(textEdit43.Text).ToString("yyyy/MM/dd");
                        nhanmauObj.Nguoinhanmau = textEdit28.Text;
                        nhanmauObj.Tinhtrangmau = textEdit39.Text;
                        nhanmauObj.Ttkhcungcap = textEdit42.Text;
                        nhanmauObj.Seal = textEdit37.Text;
                        nhanmauObj.TempMaMau = nhanmauMod.getTempMaMau();

                        for (int i = 0; i < k.Length; i++)
                        {
                            DataRow row1 = gridView58.GetDataRow(k[i]);
                            nhanmauObj.Mact = row1["MACT"].ToString();
                            nhanmauObj.Tenct = row1["TENCT"].ToString();
                            nhanmauObj.Macd = row1["MACD"].ToString();
                            nhanmauObj.Tencd = row1["TENCD"].ToString();
                            nhanmauObj.Manenmau = row1["MANENMAU"].ToString();
                            nhanmauObj.Tennenmau = row1["TENNENMAU"].ToString();
                            nhanmauObj.Madv = row1["MADV"].ToString();
                            nhanmauObj.Tendv = row1["TENDV"].ToString();
                            nhanmauObj.Lod = row1["LOD"].ToString();
                            nhanmauObj.Mapp = row1["MAPP"].ToString();
                            nhanmauObj.Tenpp = row1["TENPP"].ToString();
                            nhanmauObj.Dongia = row1["DONGIA"].ToString();
                            nhanmauObj.Nhomct = row1["NHOMCHITIEU"].ToString();
                            nhanmauObj.Maqc = row1["MAQC"].ToString();
                            nhanmauObj.Tenqc = row1["TENQC"].ToString();
                            nbjnMod.addData(nhanmauObj);
                        }
                        XtraMessageBox.Show("Nhân bản thành công");
                        loadNhanBanMau();
                        dtNhanBan2.Rows.Add(textEdit32.Text, nhanmauObj.Mahd, nhanmauObj.Jobno, nhanmauObj.Mamau, nhanmauObj.Tenmau, nhanmauObj.Time_nhanmaudk, nhanmauObj.Motamau, nhanmauObj.Nguonmau, nhanmauObj.Khoiluongmau, nhanmauObj.Tinhtrangmau, nhanmauObj.Nguoinhanmau, nhanmauObj.Seal, nhanmauObj.Time_trabcpt_dk, nhanmauObj.Ttkhcungcap, nhanmauObj.TempMaMau);
                        gridControl31.DataSource = dtNhanBan2;

                    }
                    catch (Exception ex)
                    {
                        XtraMessageBox.Show(ex.ToString());
                    }
                }


            }
        }

        private void repositoryItemButtonEdit6_Click(object sender, EventArgs e)
        {
            string message = "Bạn chắc chắn muốn xóa";
            string caption = "Lỗi cmnr";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, caption, buttons);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                int[] k = gridView56.GetSelectedRows();
                try
                {
                    //string mahd = textBoxDSMau.Text;
                    DataRow row = gridView56.GetDataRow(k[0]);
                    pycptObj.Jobno = row["JOBNO"].ToString();
                    pycptObj.Tenmau = row["TENMAU"].ToString();
                    pycptObj.Mahd = row["MAHD"].ToString();
                    nbjnMod.delMau(row["temp_MAMAU"].ToString());
                    gridView56.DeleteRow(k[0]);
                    int ma = 1;
                    foreach (DataRow row1 in njnMod.GetJobNo(pycptObj.Jobno).Rows)
                    {
                        pycptObj.TempMaMau = row1[0].ToString();
                        pycptObj.Mamau = pycptObj.Jobno + "/" + ma;
                        ma++;
                        njnMod.CapNhatJobNo(pycptObj);
                        for (int i = 0; i < gridView56.RowCount; i++)
                        {
                            DataRow row2 = gridView56.GetDataRow(i);
                            if (row2["temp_MAMAU"].ToString() == pycptObj.TempMaMau)
                            {
                                gridView56.GetDataRow(i).SetField("MAMAU", pycptObj.Mamau);
                                break;
                            }
                        }

                    }
                    loadNhanBanMau();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.ToString());
                }
            }
        }
        #endregion

        #region THỐNG KÊ CÔNG NỢ CHƯA CÓ KẾT QUẢ
        private string thoigiankt_24, thoigianbd_24;

        private void simpleButton10_Click_1(object sender, EventArgs e)
        {
            thoigianbd_24 = Convert.ToDateTime(dteNgayBD_24.EditValue).ToString("yyyy/MM/dd");
            thoigiankt_24 = Convert.ToDateTime(dteNgayKT_24.EditValue).ToString("yyyy/MM/dd");
            dgvCongNoK.DataSource = cnk_Mod.GetData(nhanmauObj, thoigianbd_24, thoigiankt_24);
            if (gridView49.RowCount == 0)
                XtraMessageBox.Show("Không có công nợ trong khoảng thời gian này");

        }
        void loadCongNoK()
        {
            cboKhachHang_24.Properties.DataSource = cnk_Mod.GetKH();
            cboKhachHang_24.Properties.ValueMember = "MAKH";
            cboKhachHang_24.Properties.DisplayMember = "TENKH";
            cboKhachHang_24.EditValue = null;

            cboJobNo_24.Properties.DataSource = cnk_Mod.GetJobNo();
            cboJobNo_24.Properties.ValueMember = "JOBNO";
            cboJobNo_24.Properties.DisplayMember = "JOBNO";
            cboJobNo_24.EditValue = null;

            cboHopDong_24.Properties.DataSource = cnk_Mod.GetHD();
            cboHopDong_24.Properties.ValueMember = "MAHD";
            cboHopDong_24.Properties.DisplayMember = "TENHD";
            cboHopDong_24.EditValue = null;

            dteNgayBD_24.EditValue = DateTime.Today;
            dteNgayKT_24.EditValue = DateTime.Today;
            dgvCongNoK.DataSource = null;
        }

        private void cboKhachHang_24_EditValueChanged(object sender, EventArgs e)
        {

            if (cboKhachHang_24.Text != "")
            {
                nhanmauObj.Makh = cboKhachHang_24.EditValue.ToString();
                cboJobNo_24.Properties.DataSource = cnk_Mod.GetJobNo(cboKhachHang_24.EditValue.ToString());
                cboJobNo_24.Properties.ValueMember = "JOBNO";
                cboJobNo_24.Properties.DisplayMember = "JOBNO";

                cboHopDong_24.Properties.DataSource = cnk_Mod.GetHD(cboKhachHang_24.EditValue.ToString());
                cboHopDong_24.Properties.ValueMember = "MAHD";
                cboHopDong_24.Properties.DisplayMember = "TENHD";
                cnObj.Makh = cboKhachHang_24.EditValue.ToString();
            }
            else
            {
                cboKhachHang_24.Text = "";
                nhanmauObj.Makh = null;

                cboKhachHang_24.Properties.DataSource = cnk_Mod.GetKH();
                cboKhachHang_24.Properties.ValueMember = "MAKH";
                cboKhachHang_24.Properties.DisplayMember = "TENKH";

                cboJobNo_24.Properties.DataSource = cnk_Mod.GetJobNo();
                cboJobNo_24.Properties.ValueMember = "JOBNO";
                cboJobNo_24.Properties.DisplayMember = "JOBNO";

                cboHopDong_24.Properties.DataSource = cnk_Mod.GetHD();
                cboHopDong_24.Properties.ValueMember = "MAHD";
                cboHopDong_24.Properties.DisplayMember = "TENHD";
            }

        }

        private void cboHopDong_24_EditValueChanged(object sender, EventArgs e)
        {
            if (cboHopDong_24.Text != "")
            {
                nhanmauObj.Mahd = cboHopDong_24.EditValue.ToString();

                cboJobNo_24.Properties.DataSource = cnk_Mod.GetJobNo1(cboHopDong_24.EditValue.ToString());
                cboJobNo_24.Properties.ValueMember = "JOBNO";
                cboJobNo_24.Properties.DisplayMember = "JOBNO";

                cboKhachHang_24.DataBindings.Clear();
                cboKhachHang_24.DataBindings.Add("EditValue", cnk_Mod.GetKH(cboHopDong_24.EditValue.ToString()), "MAKH");
            }
            else
            {
                cboHopDong_24.Text = "";
                nhanmauObj.Mahd = null;

                cboJobNo_24.Properties.DataSource = cnk_Mod.GetJobNo();
                cboJobNo_24.Properties.ValueMember = "JOBNO";
                cboJobNo_24.Properties.DisplayMember = "JOBNO";

                cboHopDong_24.Properties.DataSource = cnk_Mod.GetHD();
                cboHopDong_24.Properties.ValueMember = "MAHD";
                cboHopDong_24.Properties.DisplayMember = "TENHD";

                cboKhachHang_24.Properties.DataSource = cnk_Mod.GetKH();
                cboKhachHang_24.Properties.ValueMember = "MAKH";
                cboKhachHang_24.Properties.DisplayMember = "TENKH";

            }
        }

        private void cboJobNo_24_EditValueChanged(object sender, EventArgs e)
        {


            if (cboJobNo_24.Text != "")
            {
                nhanmauObj.Jobno = cboJobNo_24.EditValue.ToString();

                cboHopDong_24.DataBindings.Clear();
                cboHopDong_24.DataBindings.Add("EditValue", cnk_Mod.GetHD1(cboJobNo_24.EditValue.ToString()), "MAHD");

                cboKhachHang_24.DataBindings.Clear();
                cboKhachHang_24.DataBindings.Add("EditValue", cnk_Mod.GetKH1(cboJobNo_24.EditValue.ToString()), "MAKH");
            }
            else
            {
                cboJobNo_24.Text = "";
                nhanmauObj.Jobno = null;

                cboJobNo_24.Properties.DataSource = cnk_Mod.GetJobNo();
                cboJobNo_24.Properties.ValueMember = "JOBNO";
                cboJobNo_24.Properties.DisplayMember = "JOBNO";

                cboHopDong_24.Properties.DataSource = cnk_Mod.GetHD();
                cboHopDong_24.Properties.ValueMember = "MAHD";
                cboHopDong_24.Properties.DisplayMember = "TENHD";

                cboKhachHang_24.Properties.DataSource = cnk_Mod.GetKH();
                cboKhachHang_24.Properties.ValueMember = "MAKH";
                cboKhachHang_24.Properties.DisplayMember = "TENKH";



            }
        }

        private void dgvCongNoK_MouseClick(object sender, MouseEventArgs e)
        {

            if (e.Button == MouseButtons.Right)
            {
                ContextMenu m = new ContextMenu();
                System.Windows.Forms.MenuItem addDevice = new System.Windows.Forms.MenuItem("Xuất công nợ khách hàng");
                m.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { addDevice });
                dgvCongNo.ContextMenu = m;
                m.Show((Control)(sender), e.Location);

                addDevice.Click += new EventHandler(CongNoK_Click);

            }
        }
        private void CongNoK_Click(object sender, EventArgs e)
        {
            if (cboKhachHang_24.Text.Equals(""))
            {
                MessageBox.Show("Chưa chọn khách hàng");
            }
            else
            {
                thoigianbd_24 = Convert.ToDateTime(dteNgayBD_24.EditValue).ToString("yyyy/MM/dd");
                thoigiankt_24 = Convert.ToDateTime(dteNgayKT_24.EditValue).ToString("yyyy/MM/dd");
                dsct1 = cnk_Mod.GetCT(nhanmauObj, thoigianbd_24, thoigiankt_24);
                count1 = dsct1.Rows.Count;
                dsmk1 = cnk_Mod.GetTM(nhanmauObj, thoigianbd_24, thoigiankt_24);
                if (count1 == 0)
                {
                    XtraMessageBox.Show("Không có thông tin để xuất");
                }
                else
                {
                    VAT v = new VAT();
                    v.ShowDialog();
                    xuatphieucongno_K();


                }
            }

        }

        private void xuatphieucongno_K()
        {
            Microsoft.Office.Interop.Excel.Application xcelAll = new Microsoft.Office.Interop.Excel.Application();
            xcelAll.Application.Workbooks.Add(Type.Missing);

            xcelAll.Cells[1, 1] = "CÔNG TY TNHH CÔNG NGHỆ NHONHO";
            xcelAll.Cells[1, 1].Font.Bold = true;
            xcelAll.Cells[1, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[1, 1].ColumnWidth = 4;
            xcelAll.Cells[1, 2].ColumnWidth = 15;
            xcelAll.Cells[1, 3].ColumnWidth = 12;
            xcelAll.Cells[1, 4].ColumnWidth = 30;

            xcelAll.Cells[3, 1] = "Số: 03.06.17";
            xcelAll.Cells[3, 1].Font.Bold = true;
            xcelAll.Cells[3, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[4, 1] = "PHIẾU CÔNG NỢ";
            xcelAll.Cells[4, 1].Font.Bold = true;
            xcelAll.Cells[4, 1].Font.Size = 16;
            xcelAll.Cells[4, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[5, 1] = "Tên khách hàng:";
            xcelAll.Cells[5, 1].Font.Bold = true;
            xcelAll.Cells[5, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[6, 1] = "Địa chỉ:";
            xcelAll.Cells[6, 1].Font.Bold = true;
            xcelAll.Cells[6, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[7, 1] = "Thời gian:";
            xcelAll.Cells[7, 1].Font.Bold = true;
            xcelAll.Cells[7, 1].Font.Name = "Times New Roman";

            string tenkh = cboKhachHang_24.Text;
            xcelAll.Cells[5, 4] = tenkh;
            xcelAll.Cells[5, 4].Font.Bold = true;
            xcelAll.Cells[5, 4].Font.Name = "Times New Roman";

            string dckh = cnk_Mod.GetDCKH(cboKhachHang_24.Text).Rows[0]["DIACHI"].ToString();
            xcelAll.Cells[6, 4] = dckh;
            xcelAll.Cells[6, 4].Font.Bold = true;
            xcelAll.Cells[6, 4].Font.Name = "Times New Roman";

            xcelAll.Cells[7, 4] = "Từ ngày " + dteNgayBD_24.Text + " đến ngày " + dteNgayKT_24.Text;
            xcelAll.Cells[7, 4].Font.Bold = true;
            xcelAll.Cells[7, 4].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 1] = "STT";
            xcelAll.Cells[8, 1].Font.Bold = true;
            xcelAll.Cells[8, 1].RowHeight = 100;
            xcelAll.Cells[8, 1].VerticalAlignment = 2;
            xcelAll.Cells[8, 1].HorizontalAlignment = 3;
            xcelAll.Cells[8, 1].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 2] = "Ngày nhận mẫu";
            xcelAll.Cells[8, 2].Font.Bold = true;
            xcelAll.Cells[8, 2].VerticalAlignment = 2;
            xcelAll.Cells[8, 2].HorizontalAlignment = 3;
            xcelAll.Cells[8, 2].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 3] = "Mã số mẫu";
            xcelAll.Cells[8, 3].Font.Bold = true;
            xcelAll.Cells[8, 3].VerticalAlignment = 2;
            xcelAll.Cells[8, 3].HorizontalAlignment = 3;
            xcelAll.Cells[8, 3].Font.Name = "Times New Roman";

            xcelAll.Cells[8, 4] = "Mô tả mẫu";
            xcelAll.Cells[8, 4].Font.Bold = true;
            xcelAll.Cells[8, 4].VerticalAlignment = 2;
            xcelAll.Cells[8, 4].HorizontalAlignment = 3;
            xcelAll.Cells[8, 4].Font.Name = "Times New Roman";

            xcelAll.get_Range("a8", "a9").Merge(false);
            xcelAll.get_Range("b8", "b9").Merge(false);
            xcelAll.get_Range("c8", "c9").Merge(false);
            xcelAll.get_Range("d8", "d9").Merge(false);

            for (int i = 0; i < dsmk1.Rows.Count; i++)
            {
                string mamau = dsmk1.Rows[i]["MAMAU"].ToString();
                string tenmau = dsmk1.Rows[i]["TENMAU"].ToString();
                cnObj.Tenmau = tenmau;
                cnObj.Mamau = mamau;
                xcelAll.Cells[i + 10, 1] = i + 1;

                xcelAll.Cells[i + 10, 2] = DateTime.Today;

                xcelAll.Cells[i + 10, 3] = mamau;

                xcelAll.Cells[i + 10, 4] = tenmau;
                xcelAll.Cells[i + 10, 4].Font.Name = "Times New Roman";

                for (int j = 0; j <= dsct1.Rows.Count; j++)
                {
                    if (j == count1)
                    {
                        Microsoft.Office.Interop.Excel.Range t1 = xcelAll.Cells[8, j + 5];
                        Microsoft.Office.Interop.Excel.Range t2 = xcelAll.Cells[9, j + 5];
                        Microsoft.Office.Interop.Excel.Range t3 = xcelAll.Cells[4, j + 5];

                        string thanhtien = cnk_Mod.GetTT(nhanmauObj, thoigianbd_24, thoigiankt_24).Rows[0]["TONGTIEN"].ToString();
                        xcelAll.Cells[8, j + 5] = "Thành tiền";
                        xcelAll.Cells[8, j + 5].VerticalAlignment = 2;
                        xcelAll.Cells[8, j + 5].HorizontalAlignment = 3;
                        xcelAll.Cells[8, j + 5].ColumnWidth = 15;
                        xcelAll.Cells[8, j + 5].Font.Name = "Times New Roman";

                        xcelAll.Cells[i + 10, j + 5] = thanhtien;
                        xcelAll.Cells[i + 10, j + 5].Font.Name = "Times New Roman";
                        xcelAll.Cells[8, j + 5].Font.Bold = true;
                        xcelAll.get_Range(t1, t2).Merge(false);


                        xcelAll.get_Range("a4", t3).Merge(false);
                        xcelAll.Cells[4, j + 5].HorizontalAlignment = 3;
                        xcelAll.Cells[4, j + 5].VerticalAlignment = 3;
                    }
                    else
                    {
                        string tenct = dsct1.Rows[j]["TENCT"].ToString();
                        nhanmauObj.Mact = dsct1.Rows[j]["MACT"].ToString();

                        string dongia = dsct1.Rows[j]["DONGIA"].ToString();

                        xcelAll.Cells[8, j + 5].ShrinkToFit = true;
                        xcelAll.Cells[8, j + 5] = tenct;
                        xcelAll.Cells[8, j + 5].Orientation = 90;
                        xcelAll.Cells[8, j + 5].Font.Name = "Times New Roman";

                        xcelAll.Cells[9, j + 5].ShrinkToFit = true;
                        xcelAll.Cells[9, j + 5] = dongia;
                        xcelAll.Cells[9, j + 5].Orientation = 90;
                        xcelAll.Cells[9, j + 5].Font.Name = "Times New Roman";

                        nhanmauObj.Mamau = mamau;
                        nhanmauObj.Dongia = dongia;
                        System.Data.DataTable sl = cnk_Mod.GetSL(nhanmauObj, thoigianbd_24, thoigiankt_24);
                        int countSL = sl.Rows.Count;
                        string soluong;
                        if (countSL == 0)
                        {
                            soluong = "";
                            xcelAll.Cells[i + 10, j + 5].Font.Name = "Times New Roman";
                        }
                        else
                        {
                            soluong = sl.Rows[0]["SOLUONGKIEM"].ToString();
                            xcelAll.Cells[i + 10, j + 5] = soluong;
                            xcelAll.Cells[i + 10, j + 5].Font.Name = "Times New Roman";
                        }
                    }
                }
            }
            Microsoft.Office.Interop.Excel.Range c1 = xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5];
            Microsoft.Office.Interop.Excel.Range tRange = xcelAll.get_Range("a8", c1);
            tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            Microsoft.Office.Interop.Excel.Range c2 = xcelAll.Cells[dsmk1.Rows.Count + 10, 1];
            Microsoft.Office.Interop.Excel.Range c3 = xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 4];
            xcelAll.get_Range(c2, c3).Merge(false);
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1] = "Tổng phí phân tích";
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].HorizontalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].VerticalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 10, 1].Font.Name = "Times New Roman";

            string tongtien = cnk_Mod.GetTongTien(nhanmauObj, thoigianbd_24, thoigiankt_24).Rows[0]["TONGTIEN"].ToString();

            xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 5] = tongtien;
            xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 5].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 10, count1 + 5].Font.Name = "Times New Roman";

            Microsoft.Office.Interop.Excel.Range c4 = xcelAll.Cells[dsmk1.Rows.Count + 11, 1];
            Microsoft.Office.Interop.Excel.Range c5 = xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 4];
            xcelAll.get_Range(c4, c5).Merge(false);
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1] = "VAT " + (vat1 * 100) + "%";
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].HorizontalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].VerticalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 11, 1].Font.Name = "Times New Roman";


            double vat = double.Parse(tongtien) * vat1;
            xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 5] = vat;
            xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 5].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 11, count1 + 5].Font.Name = "Times New Roman";

            Microsoft.Office.Interop.Excel.Range c6 = xcelAll.Cells[dsmk1.Rows.Count + 12, 1];
            Microsoft.Office.Interop.Excel.Range c7 = xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 4];
            xcelAll.get_Range(c6, c7).Merge(false);
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1] = "Tổng tiền thanh toán";
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].HorizontalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].VerticalAlignment = 3;
            xcelAll.Cells[dsmk1.Rows.Count + 12, 1].Font.Name = "Times New Roman";

            double thanhtoan = double.Parse(tongtien) + vat;
            xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5] = thanhtoan;
            xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 12, count1 + 5].Font.Name = "Times New Roman";

            xcelAll.Cells[dsmk1.Rows.Count + 13, 1] = "Bằng chữ:";
            xcelAll.Cells[dsmk1.Rows.Count + 13, 1].Font.Name = "Times New Roman";
            xcelAll.Cells[dsmk1.Rows.Count + 15, 1] = "Lập phiếu";
            xcelAll.Cells[dsmk1.Rows.Count + 15, 1].Font.Name = "Times New Roman";
            xcelAll.Cells[dsmk1.Rows.Count + 15, 1].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 15, count1 + 4] = "Ngày " + DateTime.Today.Day + " tháng " + DateTime.Today.Month + " năm " + DateTime.Today.Year;
            xcelAll.Cells[dsmk1.Rows.Count + 15, count1 + 4].Font.Bold = true;
            xcelAll.Cells[dsmk1.Rows.Count + 15, count1 + 4].Font.Name = "Times New Roman";

            xcelAll.Visible = true;
        }
        #endregion

        #region Thống kê chỉ tiêu
        void loadThongKeChiTieu()
        {
            dateTimePicker1.EditValue = DateTime.Now;
            dateTimePicker2.EditValue = DateTime.Now;
            dgvThongKeChiTieu.DataSource = null;
        }
        private void simpleButton79_Click(object sender, EventArgs e)
        {
            string tu = "", den = "";
            if (dateTimePicker1.EditValue != null)
                tu = Convert.ToDateTime(dateTimePicker1.EditValue).ToString("yyyy/MM/dd");
            if (dateTimePicker2.EditValue != null)
                den = Convert.ToDateTime(dateTimePicker2.EditValue).ToString("yyyy/MM/dd");
            if (tu == "" || den == "")
                XtraMessageBox.Show("Vui lòng chọn khoản thời gian cần thống kê");
            else
                dgvThongKeChiTieu.DataSource = tkMod.GetData_ThongKeChiTieu(tu, den);
            if (gridView62.RowCount == 0)
                XtraMessageBox.Show("Không có chỉ tiêu kiểm trong khoảng thời gian này!");

        }

        // xUÁT EXCEL
        private void simpleButton80_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    XtraMessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "THỐNG KÊ CHỈ TIÊU";
                if (ws == null)
                {
                    XtraMessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                int fontSizeTenTruong = 14;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A2");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 15;
                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B2");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN CHỈ TIÊU";
                row23_MaSP.ColumnWidth = 70;

                // Cột 3: tÊN NỀN MẪU
                Microsoft.Office.Interop.Excel.Range row23_NM = ws.get_Range("C2");//Cột B dòng 2 và dòng 3
                row23_NM.Merge();
                row23_NM.Font.Size = fontSizeTenTruong;
                row23_NM.Font.Name = fontName;
                row23_NM.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_NM.Value2 = "NỀN MẪU";
                row23_NM.ColumnWidth = 30;

                // Cột 3: Số lần kiểm
                Microsoft.Office.Interop.Excel.Range row23_SLK = ws.get_Range("D2");//Cột B dòng 2 và dòng 3
                row23_SLK.Merge();
                row23_SLK.Font.Size = fontSizeTenTruong;
                row23_SLK.Font.Name = fontName;
                row23_SLK.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_SLK.Value2 = "SỐ LẦN KIỂM";
                row23_SLK.ColumnWidth = 20;


                string tenct;
                string slk;
                string nm;
                int k = gridView62.DataRowCount;
                for (int i = 0; i < k; i++)
                {
                    DataRow row = gridView62.GetDataRow(i);
                    tenct = row["TENCT"].ToString();
                    nm = row["TENNENMAU"].ToString();
                    slk = row["SOLANKIEM"].ToString();

                    Microsoft.Office.Interop.Excel.Range row23_sttct = ws.get_Range("A" + (i + 3));
                    row23_sttct.Value2 = i + 1;
                    Microsoft.Office.Interop.Excel.Range row23_tenct = ws.get_Range("B" + (i + 3));
                    row23_tenct.Value2 = tenct;
                    Microsoft.Office.Interop.Excel.Range row23_nm = ws.get_Range("C" + (i + 3));
                    row23_nm.Value2 = nm;
                    Microsoft.Office.Interop.Excel.Range row23_slk = ws.get_Range("D" + (i + 3));
                    row23_slk.Value2 = slk;
                }


                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A2", "D2");
                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //in đậm
                row23_CotTieuDe.Font.Bold = true;
                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A2", "D2"));
                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

        }

        #endregion
       
        #region QUẢN LÝ CHỈ ĐỊNH
        void enableChiDinh(bool e)
        {
            txtTenChiDinh.ReadOnly = e;
            dgvChiDinh.Enabled = e;
            gpCTTCD.Enabled = e;
            if (e)
            {
                txtTenChiDinh.BackColor = Color.Linen;
                layoutControlItem76.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem289.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem79.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            }

            else
            {
                txtTenChiDinh.BackColor = Color.White;
                layoutControlItem76.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem289.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlItem79.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }
        }
        private void simpleButton11_Click(object sender, EventArgs e)
        {
            enableChiDinh(false);
            dgvChiTieu2.DataSource = cdMod.GetDataChiTieuNenMau();
            gpThemCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            gpTTCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            simpleButton11.Visible = false;
            btnCapNhatCD.Visible = false;
           
            dgvChiDinh.Enabled = false;
            btnXoaCD.Visible = false;
           
            loadThemCDMoi();
        }

        private void btnCapNhatCD_Click(object sender, EventArgs e)
        {

            if (txtTenChiDinh.Text == "")
            {
                labelControl1.Text = "Tên chỉ định không được bỏ trống!";
                return;
            }
            else
            {
                Chi_Dinh_Obj cdObj = new Chi_Dinh_Obj();
                cdObj.Macd = int.Parse(txtMaChiDinh.Text.Trim());
                cdObj.Tencd = txtTenChiDinh.Text.Trim();
                if (cdMod.KiemTraTenCHiDinh(txtTenChiDinh.Text.Trim(), txtMaChiDinh.Text.Trim()).Rows.Count > 0)
                {
                    labelControl1.Text = "Tên chỉ định đã tồn tại vui lòng điền tên khác!";
                    return;
                }
                else
                {
                    labelControl1.Text = "";
                    cdMod.update(cdObj);
                }

                loadChiDinh();
            }
        }

        private void btnNapLaiChiDinh_Click(object sender, EventArgs e)
        {
            txtMaChiDinh_EditValueChanged(sender, e);
            loadChiDinh();

            gpDSCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gpCTTCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            gpThemCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gpTTCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;

        }

        private void simpleButton45_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls|Excel Workbook|*.xlsm", ValidateNames = true };
            textBox4.Text = ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : "";
        }

        private void simpleButton43_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    XtraMessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;

                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "Chỉ Định";
                if (ws == null)
                {
                    XtraMessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //int row = 1;
                string fontName = "Times New Roman";
                int fontSizeTenTruong = 14;

                // Cột 1:  Tạo Ô Số Thứ Tự (STT)
                Microsoft.Office.Interop.Excel.Range row23_STT = ws.get_Range("A1");//Cột A dòng 2 và dòng 3
                row23_STT.Merge();
                row23_STT.Font.Size = fontSizeTenTruong;
                row23_STT.Font.Name = fontName;
                row23_STT.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_STT.Value2 = "STT";
                row23_STT.ColumnWidth = 5.5;

                // Cột 2: Tạo Ô Ngày Nhận Mẫu :
                Microsoft.Office.Interop.Excel.Range row23_MaSP = ws.get_Range("B1");//Cột B dòng 2 và dòng 3
                row23_MaSP.Merge();
                row23_MaSP.Font.Size = fontSizeTenTruong;
                row23_MaSP.Font.Name = fontName;
                row23_MaSP.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row23_MaSP.Value2 = "TÊN CHỈ ĐỊNH";
                row23_MaSP.ColumnWidth = 50;

                //Tô nền vàng các cột tiêu đề:
                Microsoft.Office.Interop.Excel.Range row23_CotTieuDe = ws.get_Range("A1", "B1");

                //nền vàng
                row23_CotTieuDe.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                //in đậm
                row23_CotTieuDe.Font.Bold = true;

                //chữ đen
                row23_CotTieuDe.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);

                //Kẻ khung toàn bộ
                BorderAround(ws.get_Range("A1", "B1"));

                // Mở chương trình Excel
                xlApp.Visible = true;

                //thoát và thu hồi bộ nhớ cho COM
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

        }

        private void simpleButton44_Click(object sender, EventArgs e)
        {

            if (textBox4.Text == "" || textBox4.Text == "Chưa có tệp nào được chọn")
            {
                XtraMessageBox.Show("Vui lòng nhập đường dẫn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox4.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] a = new string[colCount];
            int sl = 0, sl2 = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 == null)
                    {
                        a[j - 2] = "";
                    }
                    else
                    {
                        a[j - 2] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                cdObj.Tencd = a[0];

                if (!cdMod.TrungTenCD(cdObj))
                {
                    cdMod.addData(cdObj);
                    sl++;
                }
                else
                {
                    sl2++;
                }
            }
            XtraMessageBox.Show("Đã thêm " + sl + " đơn vị vào CSDL!  --Có " + sl2 + " chỉ định đã tồn tại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.None);
            loadChiDinh();
            textBox4.Text = "Chưa có tệp nào được chọn";
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

        }

        private void loadChiDinh()
        {
            enableChiDinh(true);
            dgvChiDinh.DataSource = cdMod.GetDataChiDinh(null);
            dgvChiDinh.Enabled = true;
            txtMaChiDinh.DataBindings.Clear();
            txtMaChiDinh.DataBindings.Add("Text", dgvChiDinh.DataSource, "MACD");
            txtTenChiDinh.DataBindings.Clear();
            txtTenChiDinh.DataBindings.Add("Text", dgvChiDinh.DataSource, "TENCD");
            //btnLuuCT.Visible = false;
            btnCapNhatCD.Visible = true;
            txtTenChiDinh.Enabled = true;
            simpleButton11.Visible = true;
            btnXoaCD.Hide();

            lblErrorCD2.Text = "";
            layoutControlGroup17.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            emptySpaceItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void loadThemCDMoi()
        {
            ErrorThemCD.Text = "";
            txtThemMoiCD.Text = "";
        
        }

        

        private void simpleButton12_Click(object sender, EventArgs e)
        {
            lblErrorCD2.Text = "";
            btnCapNhatCD.Visible = false;
            //btnLuuCT.Visible = true;
            btnXoaCD.Visible = false;
            gpDSCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            gpCTTCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            enableChiDinh(true);
            dgvChiDinh.Enabled = false;

        }

        private void simpleButton41_Click(object sender, EventArgs e)
        {
            //btnCapNhatCD.Visible = true;
            //btnLuuCT.Visible = false;
            //gpDSCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            //gpCTTCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            ///
            Chi_Dinh_Chi_Tieu_Obj cdctObj = new Chi_Dinh_Chi_Tieu_Obj();
            int[] k = gridView11.GetSelectedRows();
            cdctObj.Macd = int.Parse(txtMaChiDinh.Text.Trim());
            cdctObj.Tencd = txtTenChiDinh.Text.Trim();
            for (int i = 0; i < k.Length; i++)
            {
                DataRow row1 = gridView11.GetDataRow(k[i]);
                cdctObj.Mact = row1[0].ToString();
                cdctObj.Tenct = row1[1].ToString();
                cdMod.addDataChiDinhCT(cdctObj);
            }
            dgvCTthuocCD.DataSource = cdMod.getDataChiDinhChiTieu(txtMaChiDinh.Text);
            dgvChiTieu1.DataSource = cdMod.getDataChiTieu(txtMaChiDinh.Text);
        }

        private void btnXoaCD_Click(object sender, EventArgs e)
        {
            int[] k = gridView13.GetSelectedRows();
            if (k.Length == 0)
            {
                labelControl1.Text = "Chưa chọn chỉ tiêu";
                return;
            }
            else
            {
                if (k.Length == gridView13.RowCount)
                {
                    labelControl1.Text = "Không thể xóa tất cả chỉ tiêu, chỉ định phải có ích nhất 1 chỉ tiêu!";
                    return;
                }
                Chi_Dinh_Chi_Tieu_Obj cdct = new Chi_Dinh_Chi_Tieu_Obj();
                if (XtraMessageBox.Show("Bạn có chắn chắn xóa?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.No)
                {
                    for (int i = 0; i < k.Length; i++)
                    {
                        DataRow row = gridView13.GetDataRow(k[i]);
                        cdct.Macd = int.Parse(row[0].ToString());
                        cdct.Mact = row[1].ToString();
                        cdMod.delectDataCTCD(cdct);
                    }
                    txtMaChiDinh_EditValueChanged(sender, e);
                }
                else
                    return;
            }
        }

        private void simpleButton28_Click(object sender, EventArgs e)
        {
          
        
            if (txtThemMoiCD.Text == "")
            {
                ErrorThemCD.Text = "Chưa điền tên chỉ định ";
                return;
            }
            else
            {
                int[] k = gridView14.GetSelectedRows();
                if (k.Length == 0)
                {
                    ErrorThemCD.Text = "Vui lòng chọn ít nhất 1 chỉ tiêu cho chỉ định";
                    return;
                }
                else
                {
                    if (cdMod.KiemTraTenCHiDinh(txtThemMoiCD.Text.Trim(), null).Rows.Count > 0)
                    {
                        ErrorThemCD.Text = "Tên chỉ định đã tồn tại vui lòng điền tên chỉ định mới";
                        return;
                    }
                    else
                    {
                        cdMod.addDataChiDinh(txtThemMoiCD.Text.Trim());
                        System.Data.DataTable dtThemCD = new System.Data.DataTable();
                        dtThemCD = cdMod.KiemTraTenCHiDinh(txtThemMoiCD.Text.Trim(), null);
                        Chi_Dinh_Chi_Tieu_Obj cdctObj = new Chi_Dinh_Chi_Tieu_Obj();
                        DataRow row = dtThemCD.Rows[0];
                        cdctObj.Macd = int.Parse(row[0].ToString());
                        cdctObj.Tencd = row[1].ToString();
                        for (int i = 0; i < k.Length; i++)
                        {
                            DataRow row1 = gridView14.GetDataRow(k[i]);
                            cdctObj.Mact = row1[0].ToString();
                            cdctObj.Tenct = row1[1].ToString();
                            cdMod.addDataChiDinhCT(cdctObj);
                        }
                        btnNapLaiChiDinh_Click(sender, e);
                    }
                }
            }
        
        }

        private void simpleButton30_Click(object sender, EventArgs e)
        {
            navigationFrame1.SelectedPage = navigationPage2;
        }

        private void txtMaChiDinh_EditValueChanged(object sender, EventArgs e)
        {
            dgvCTthuocCD.DataSource = cdMod.getDataChiDinhChiTieu(txtMaChiDinh.Text);
            dgvChiTieu1.DataSource = cdMod.getDataChiTieu(txtMaChiDinh.Text);
            labelControl1.Text = "";
            layoutControlGroup17.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            emptySpaceItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void gridView13_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {

            int[] k = gridView13.GetSelectedRows();
            if (k.Length > 0)
            {
                emptySpaceItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                layoutControlGroup17.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            }
            else
            {
                layoutControlGroup17.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }
        private void capnhatChiDinh_Click(object sender, EventArgs e)
        {
            enableChiDinh(false);
        }

        private void btnxoaChiDinh_Click(object sender, EventArgs e)
        {

            if (gridView12.RowCount > 0)
            {
                int[] i = gridView12.GetSelectedRows();
                DataRow row = gridView12.GetDataRow(i[0]);
                if (cdMod.GetDataChiDinh(row[0].ToString()).Rows.Count == 0)
                {
                    cdMod.delectCD(row[0].ToString());
                    loadChiDinh();
                }
                else
                {
                    labelControl1.Text = "Không thể xóa chỉ định";
                    return;
                }
            }

        }

        private void simpleButton22_Click(object sender, EventArgs e)//hủy
        {
            txtMaChiDinh_EditValueChanged(sender, e);
            loadChiDinh();

            gpDSCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gpCTTCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            gpThemCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gpTTCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
        }

        private void simpleButton23_Click(object sender, EventArgs e)//hủy
        {
            txtMaChiDinh_EditValueChanged(sender, e);
            loadChiDinh();

            gpDSCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gpCTTCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            gpThemCD.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            gpTTCT.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
        }
        #endregion

        #region DANH SÁCH THEO DỖI MẪU,KIỂM NGHIỆM
        void loadDStheodoiMau()
        {
            gridControl6.DataSource=cnmMod.GetDSMau(false);
        }
        void loadDStheodoikiemNghiem()
        {
            gridControl10.DataSource = dkqMod.dsTheoDoiKiemNghiem();
        }
        #endregion

        #region XUẤT KẾT QUẢ(RAW)

        private void checkEdit7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit7.Checked == true)
                layoutControlItem300.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
            else
                layoutControlItem300.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void layoutControlItem298_TextChanged(object sender, EventArgs e)
        {
            gridControl32.DataSource = xkq.getChiTieuRAW(layoutControlItem298.Text);
        }

        private void simpleButton31_Click(object sender, EventArgs e)
        {
            int[] k = gridView66.GetSelectedRows();
            if (k.Length == 0)
                XtraMessageBox.Show("Chưa chọn chỉ tiêu cần xuất");
            else
            {
                tableXuatKQ();
                for (int i = 0; i < k.Length; i++)
                {
                    DataRow row1 = gridView66.GetDataRow(k[i]);
                    cnObj.Maphieu = row1["MAPHIEU_YCPT"].ToString();
                    cnObj.Mahd = row1["MAHD"].ToString();
                    cnObj.Jobno = row1["JOBNO"].ToString();
                    cnObj.Mamau = row1["MAMAU"].ToString();
                    cnObj.Tenmau = row1["TENMAU"].ToString();
                    cnObj.Time_nhanmau_tt = Convert.ToDateTime(row1["TIME_NHANMAU_TT"].ToString()).ToString("yyyy/MM/dd");
                    cnObj.Manenmau = row1["MANENMAU"].ToString();
                    if (cnObj.Manenmau == "")
                        cnObj.Manenmau = "NULL";
                    cnObj.Nenmau = row1["TENNENMAU"].ToString();
                    cnObj.Machidinh = row1["MACD"].ToString();
                    if (cnObj.Machidinh == "")
                        cnObj.Machidinh = "NULL";
                    cnObj.Chidinh = row1["TENCD"].ToString();
                    cnObj.Machitieu = row1["MACT"].ToString();
                    cnObj.Chitieu = row1["TENCTPHULUC"].ToString();
                    cnObj.Maqc = row1["MAQC"].ToString();
                    if (cnObj.Maqc == "")
                        cnObj.Maqc = "NULL";
                    cnObj.Tenqc = row1["TENQC"].ToString();
                    cnObj.Maphuongphap = row1["MAPP"].ToString();
                    if (cnObj.Maphuongphap == "")
                        cnObj.Maphuongphap = "NULL";
                    cnObj.Phuongphap = row1["TENPP"].ToString();
                    cnObj.Dongia = row1["DONGIA"].ToString();
                    cnObj.Lankiemthu = cnMod.LanKiemThu(cnObj.Maphieu);
                    cnObj.Soluong = row1["SOLUONGKIEM"].ToString();
                    if (cnObj.Soluong == "")
                        cnObj.Soluong = "1";
                    cnObj.Lod = row1["LOD"].ToString();
                    cnObj.Madv = row1["MADV"].ToString();
                    if (cnObj.Madv == "")
                        cnObj.Madv = "NULL";
                    cnObj.Tendv = row1["TENDV"].ToString();
                    cnObj.Nhomchitieu = row1["NHOMCHITIEU"].ToString();
                    cnObj.Userxuat = Login.tenUser;
                    cnObj.Makh = row1["MAKH"].ToString();
                    cnObj.Tenkh = row1["TENKH"].ToString();
                    cnObj.Ketqua = row1["KETQUA_PTPHULUC"].ToString();
                    
                    dtXuatKQ.Rows.Add(row1["JOBNO"], row1["TENKH"], row1["DIACHI"], row1["TTKHCUNGCAP"], row1["MOTAMAU"], row1["SEAL"], row1["NGUONMAU"], row1["TIME_NHANMAU_TT"], row1["TIME_CHUYENMAU"], row1["MAMAU"], row1["TENMAU"], row1["MACT"], row1["TENCT"], row1["MAPP"], row1["TENPP"], row1["MADV"], row1["TENDV"], row1["LOD"], row1["KETQUA_PT"], row1["MAQC"], row1["TENQC"], row1["NHOMCHITIEU"], textEdit10.Text);
                    if (row1["TENCT"].ToString() != row1["TENCTPHULUC"].ToString())
                    {
                        dtXuatKQ_PhuLuc.Rows.Add(row1["MAMAU"], row1["TENMAU"], row1["MACT"], row1["TENCTPHULUC"], row1["MAPP"], row1["TENPP"], row1["MADV"], row1["TENDV"], row1["LOD"], row1["KETQUA_PTPHULUC"], row1["MAQC"], row1["TENQC"], row1["NHOMCHITIEU"], searchLookUpEdit5.Text);
                    }

                }
                DataSet ds1 = new DataSet();
                ds1.Tables.Add(dtXuatKQ);
                ds1.Tables.Add(dtXuatKQ_PhuLuc);
                Report.frmrpXuatKQ.dt2 = dtXuatKQ.Copy();
                Report.rpXuatKQ.dtPhuLuc = dtXuatKQ_PhuLuc.Copy();
                Report.frmrpXuatKQ__QC.dt2 = dtXuatKQ.Copy();
                Report.rpXuatKQ_QC.dtPhuLuc = dtXuatKQ_PhuLuc.Copy();
                ds1.WriteXmlSchema("ReportBaoCaoKetQua.xml");

                if (checkEdit11.Checked == true)
                {
                    frmrpXuatKQ__QC f = new frmrpXuatKQ__QC();
                    f.ShowDialog();
                }
                else
                {
                    frmrpXuatKQ f = new frmrpXuatKQ();
                    f.ShowDialog();
                }
                loadXuatKetQua();
            }
        }
        #endregion





    }
}