using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;

namespace QuanLyMauKiemNghiem.Report
{
    public partial class rpXuatKQ : DevExpress.XtraReports.UI.XtraReport
    {
        public rpXuatKQ()
        {
            InitializeComponent();
        }
        int count = 0;
        private void PageHeader_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            
            if (count == 0)
            {
                subBand1.Visible = false;
                subBand2.Visible = true;
                PageFooter.Visible = true;
                count++;
            }
            else
            {
                subBand1.Visible = true;
                subBand2.Visible = false;
                PageFooter.Visible = false;
                if (this.RowCount == 1)
                {
                    if (dtPhuLuc.Rows.Count != 0)
                    {
                        subBand3.Visible = false;
                    }
                                    }
             
            }
        }

        private int rowCounter = 0;
        private void subBand3_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            
                if (this.RowCount > 1)
                {

                    if (rowCounter >= this.RowCount)
                    { e.Cancel = true; }
                }
            
        }

        private void Detail_AfterPrint(object sender, EventArgs e)
        {
            
            rowCounter++;
        }
        public static DataTable dtPhuLuc;
        private void xrSubreport1_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {   DataSet ds1= new DataSet();
            ds1.Tables.Add(dtPhuLuc);
            subPhuLuc rp2 = (subPhuLuc)xrSubreport1.ReportSource;
            rp2.DataSource = ds1;
        }

        private void rpXuatKQ_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            
            if (dtPhuLuc.Rows.Count == 0)
                
                SubBand4.Visible = false;
            
        }





    }
}
