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
using ZXing;
using QuanLyMauKiemNghiem.Model;


namespace QuanLyMauKiemNghiem.View
{
    public partial class QRCode : DevExpress.XtraEditors.XtraForm
    {
        QR_Mod qrMod = new QR_Mod();
        public QRCode()
        {
            InitializeComponent();
        }

        private void QRCode_FormClosing(object sender, FormClosingEventArgs e)
        {
       
            cameraControl1.Stop();
        }

        private void QRCode_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {

                BarcodeReader reader = new BarcodeReader();
                Result result = reader.Decode((Bitmap)cameraControl1.TakeSnapshot());
                if (result != null)
                {
                    string QR = result.ToString().Trim();
                    string[] mang = QR.Split(' ');
                    string decode = mang[0].ToString();
                    timer1.Stop();
                    DataTable dt = new DataTable();
                    dt = qrMod.KiemTramau(decode,false);
                    if (dt.Rows.Count > 0)
                    {
                        DataRow row = dt.Rows[0];
                        if (row[2].ToString() == "")
                        {
                            XtraMessageBox.Show("Mẫu chưa có chỉ tiêu, Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                           // FinalFrame.Stop();
                            
                            this.Close();
                            return;
                        }

                        if (row[1].ToString() != "Đã chuyển")//1= tình trạng chuyển mẫu là đã chuyển thành công
                        {
                            if (row[1].ToString() == "Hoàn thành")
                            {
                                XtraMessageBox.Show("Mẫu đã kiểm hoàn thành, Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                timer1.Start();
                                return;
                            }
                            else
                            {
                                if (row[1].ToString() == "Đã xuất")
                                {
                                    if (qrMod.KiemTramau(decode, true).Rows.Count == 0)
                                    {
                                        XtraMessageBox.Show("Mẫu đã xuất kết quả, Vui lòng kiểm tra lại!");
                                        timer1.Start();
                                        return;
                                    }                                   
                                }
                            }
                            qrMod.updateTrangThaiChuyen(decode);
                            if (XtraMessageBox.Show("Mẫu '" + row[0] + "' nhận thành công!, Bạn có muốn tiếp tục nhận mẫu mới", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) != DialogResult.Yes)
                            {                              
                                this.Close();
                            }
                            else
                                timer1.Start();
                        }
                        else
                        {
                            if (XtraMessageBox.Show("Mẫu '" + row[0] + "' đã nhận rồi!, Bạn có muốn tiếp tục nhận mẫu mới", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                            {
                              
                                this.Close();
                            }
                            else
                                timer1.Start();
                        }
                    }
                    else
                    {
                        if (XtraMessageBox.Show("Mẫu '" + decode + "' không tồn tại!, Bạn có muốn tiếp tục nhận mẫu mới", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        {
                          
                            this.Close();
                        }
                        else
                            timer1.Start();
                    }
                }

            }
            catch
            {

            }
        }
    }
}