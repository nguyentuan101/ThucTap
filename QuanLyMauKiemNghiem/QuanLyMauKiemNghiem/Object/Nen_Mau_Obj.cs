using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Nen_Mau_Obj
    {
        private string manenmau, tennenmau;

        public string Manenmau
        {
            get { return manenmau; }
            set { manenmau = value; }
        }

        public string Tennenmau
        {
            get { return tennenmau; }
            set { tennenmau = value; }
        }



        public Nen_Mau_Obj()
        {
        }
    }
}
