using DevExpress.XtraReports.UI;

namespace QuanLyMauKiemNghiem.Report
{
    public partial class rpPhieuYCPT : DevExpress.XtraReports.UI.XtraReport
    {
        public rpPhieuYCPT()
        {
            InitializeComponent();
        }

        //private int count = 0;

     

        private void xrTable3_PrintOnPage(object sender, PrintOnPageEventArgs e)
        {
            //if (e.PageIndex != e.PageCount - 1)
            //{
            //    e.Cancel=true;
            //}
        }

        private void xrTableCell1_PrintOnPage(object sender, PrintOnPageEventArgs e)
        {
        }

        private int rowCounter = 0;
        private void Detail_AfterPrint(object sender, System.EventArgs e)
        {
            rowCounter++;
        }
        int count = 0;

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

      
    }
}