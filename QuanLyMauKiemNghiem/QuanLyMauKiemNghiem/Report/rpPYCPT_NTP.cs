using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;

namespace QuanLyMauKiemNghiem.Report
{
    public partial class rpPYCPT_NTP : DevExpress.XtraReports.UI.XtraReport
    {
        public rpPYCPT_NTP()
        {
            InitializeComponent();
        }

        private void TopMargin_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            
        }

        private void TopMargin_AfterPrint(object sender, EventArgs e)
        {
            
        }
        private int rowCounter = 0;
        int count;
        private void PageHeader_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {

            if (count == 0)
            {
                if (this.RowCount == 1)
                {
                    e.Cancel = false;
                }
                else
                {
                    if (rowCounter > this.RowCount)
                        e.Cancel = true;
                }
                count++;
            }
            else
            {
                if (rowCounter >= this.RowCount) { e.Cancel = true; }


            }
        }

        private void Detail_AfterPrint(object sender, EventArgs e)
        {
            rowCounter++;
        }



     



    }
}
