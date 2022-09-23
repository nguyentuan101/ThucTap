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
    public partial class frmXuatPYCPT : DevExpress.XtraEditors.XtraForm
    {
        public frmXuatPYCPT()
        {
            InitializeComponent();
        }
        public static DataTable dt;
        private void frmXuatPYCPT_Load(object sender, EventArgs e)
        {

            rpPhieuYCPT rp = new rpPhieuYCPT();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);


            rp.DataSource = ds;
            documentViewer1.DocumentSource = rp;
            rp.CreateDocument();
        }
    }
}