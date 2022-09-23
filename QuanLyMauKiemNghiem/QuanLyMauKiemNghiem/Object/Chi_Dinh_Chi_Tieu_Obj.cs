using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Chi_Dinh_Chi_Tieu_Obj
    {

        int macd;

        public int Macd
        {
            get { return macd; }
            set { macd = value; }
        }
        string mact, tenct, tencd;

        public string Tencd
        {
            get { return tencd; }
            set { tencd = value; }
        }

        public string Tenct
        {
            get { return tenct; }
            set { tenct = value; }
        }

        public string Mact
        {
            get { return mact; }
            set { mact = value; }
        }

        public Chi_Dinh_Chi_Tieu_Obj()
        {
        }
    }
}
