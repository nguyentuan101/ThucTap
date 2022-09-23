using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class SD_Quyen_Obj
    {
        string tenuser, idquyen, tenquyen;

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

        public string Tenuser
        {
            get { return tenuser; }
            set { tenuser = value; }
        }



        public SD_Quyen_Obj()
        {

        }

       
    }
}
