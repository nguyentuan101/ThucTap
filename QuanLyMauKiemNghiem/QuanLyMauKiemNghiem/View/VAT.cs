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

namespace QuanLyMauKiemNghiem.View
{
    public partial class VAT : DevExpress.XtraEditors.XtraForm
    {
        public VAT()
        {
            InitializeComponent();
        }
   
        private void VAT_Load(object sender, EventArgs e)
        {
            
        }

        private string thue;
        public void simpleButton1_Click(object sender, EventArgs e)
        {
            thue = cboVAT.EditValue.ToString();
           // frmMain fmain = new frmMain();
           // this.Close();
            frmMain.vat1 = double.Parse(thue);

            this.Close();

        }
        public bool temp = false;        
        public bool tam()
        {
            temp = true;
            return temp;
        }
    }
}