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

namespace QuanLyMauKiemNghiem.Report
{
    public partial class frmXuatPYCPT_NTP : DevExpress.XtraEditors.XtraForm
    {
        public frmXuatPYCPT_NTP()
        {
            InitializeComponent();
        }
        public static DataTable dt2;
        private void frmXuatPYCPT_NTP_Load(object sender, EventArgs e)
        {
            rpPYCPT_NTP rp = new rpPYCPT_NTP();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt2);


            rp.DataSource = ds;
            documentViewer1.DocumentSource = rp;
            rp.CreateDocument();
        }
    }
}