using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
   public class Nguoi_Dung_Obj
    {
        string tenuser, nhom, hoten, matkhau, matkhaucu;

        public string Matkhaucu
        {
            get { return matkhaucu; }
            set { matkhaucu = value; }
        }
        bool xoand,truongnhom;



        public string Matkhau
        {
            get { return matkhau; }
            set { matkhau = value; }
        }
      
        public string Hoten
        {
            get { return hoten; }
            set { hoten = value; }
        }

        public string Nhom
        {
            get { return nhom; }
            set { nhom = value; }
        }

        public string Tenuser
        {
            get { return tenuser; }
            set { tenuser = value; }
        }
        

       public bool Truongnhom
        {
            get { return truongnhom; }
            set { truongnhom = value; }
        }

       public bool Xoand
        {
            get { return xoand; }
            set { xoand = value; }
        }



        public Nguoi_Dung_Obj()
        {

        }

    
    }
}
