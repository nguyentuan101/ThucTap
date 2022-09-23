using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Chi_Dinh_Obj
    {
        int macd;

        public int Macd
        {
            get { return macd; }
            set { macd = value; }
        }
        string tencd;

        public string Tencd
        {
            get { return tencd; }
            set { tencd = value; }
        }

        public Chi_Dinh_Obj() { }
    }
}
