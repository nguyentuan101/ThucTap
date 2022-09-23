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
using DevExpress.XtraReports.UI;
using QuanLyMauKiemNghiem.Model;

namespace QuanLyMauKiemNghiem.Report
{
    public partial class frmrpXuatKQ : DevExpress.XtraEditors.XtraForm
    {
        public frmrpXuatKQ()
        {
            InitializeComponent();
        }
        public static DataTable dt2;
        public static DataTable dtPhuLuc;
        LoadData ld = new LoadData();
        private void frmrpXuatKQ_Load(object sender, EventArgs e)
        {
            rpXuatKQ rp1 = new rpXuatKQ();
            
            DataSet ds = new DataSet();
            ds.Tables.Add(dt2);
            rp1.DataSource = ds;
            documentViewer1.DocumentSource = rp1;
            rp1.CreateDocument();

            //subPhuLuc rp = new subPhuLuc();
            //XtraReport1 rp = new XtraReport1();
            //DataSet ds1 = new DataSet();           
            //ds1.Tables.Add(dtPhuLuc);
            //rp.DataSource = ds1;
            //documentViewer1.DocumentSource = rp;
            //rp.CreateDocument();


        }
    }
}