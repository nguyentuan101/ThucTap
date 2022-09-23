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
    public partial class frmrpQRCode : DevExpress.XtraEditors.XtraForm
    {
        public static DataTable dt1;
        public frmrpQRCode()
        {
            InitializeComponent();
        }

        private void frmrpQRCode_Load(object sender, EventArgs e)
        {
            rpt_QR rp = new rpt_QR();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt1);


            rp.DataSource = ds;
            documentViewer1.DocumentSource = rp;
            rp.CreateDocument();
        }
    }
}