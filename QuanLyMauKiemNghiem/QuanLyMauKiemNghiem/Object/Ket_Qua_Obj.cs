using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Ket_Qua_Obj
    {
        string makq, maphieu, mapp, tenpp, ketquapt, ghichupt, userNhapKQ;
        public Ket_Qua_Obj()
        {
        }
        public string Makq
        {
            get { return makq; }
            set { makq = value; }
        }

        public string Maphieu
        {
            get { return maphieu; }
            set { maphieu = value; }
        }

        public string Mapp
        {
            get { return mapp; }
            set { mapp = value; }
        }

        public string Tenpp
        {
            get { return tenpp; }
            set { tenpp = value; }
        }

        public string Ketquapt
        {
            get { return ketquapt; }
            set { ketquapt = value; }
        }

        public string Ghichupt
        {
            get { return ghichupt; }
            set { ghichupt = value; }
        }

        public string UserNhapKQ
        {
            get { return userNhapKQ; }
            set { userNhapKQ = value; }
        }

    
    }
}
