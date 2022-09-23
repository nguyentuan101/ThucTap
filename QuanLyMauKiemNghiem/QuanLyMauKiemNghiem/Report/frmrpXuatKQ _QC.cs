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
    public partial class frmrpXuatKQ__QC : DevExpress.XtraEditors.XtraForm
    {
        public frmrpXuatKQ__QC()
        {
            InitializeComponent();
        }
        public static DataTable dt2;
       
        private void frmrpXuatKQ__QC_Load(object sender, EventArgs e)
        {
            rpXuatKQ_QC rp = new rpXuatKQ_QC();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt2);
       

            rp.DataSource = ds;
            documentViewer1.DocumentSource = rp;
            rp.CreateDocument();

        }
    }
}