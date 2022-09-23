using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    class DS_Quyen_Obj
    {
        private string idquyen, tenquyen;

        public string Tenquyen
        {
            get { return tenquyen; }
            set { tenquyen = value; }
        }

        public string Idquyen
        {
            get { return idquyen; }
            set { idquyen = value; }
        }

        public DS_Quyen_Obj()
        {
        }
    }
}
