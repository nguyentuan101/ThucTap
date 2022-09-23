using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
   public class Noi_Kiem_Obj
    {
        private string tennoikiem, loai, idkiem;

        public string Idkiem
        {
            get { return idkiem; }
            set { idkiem = value; }
        }
        
        bool xoa;


        public string Loai
        {
            get { return loai; }
            set { loai = value; }
        }
                
    
       public string Tennoikiem
       {
           get { return tennoikiem; }
           set { tennoikiem = value; }
       }

        public bool Xoa
        {
            get { return xoa; }
            set { xoa = value; }
        }


        public Noi_Kiem_Obj()
        { }


    }
}
