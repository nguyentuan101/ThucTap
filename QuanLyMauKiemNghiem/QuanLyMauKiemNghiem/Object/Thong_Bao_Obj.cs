using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Thong_Bao_Obj
    {
        int idthongbao;
        string noidung, tenuser;

        public int Idthongbao
        {
            get { return idthongbao; }
            set { idthongbao = value; }
        }
        

        public string Tenuser
        {
            get { return tenuser; }
            set { tenuser = value; }
        }

        public string Noidung
        {
            get { return noidung; }
            set { noidung = value; }
        }

        public Thong_Bao_Obj()
        {

        }

        public Thong_Bao_Obj(int idthongbao, string noidung, string tenuser)
        {
            idthongbao = 1;
            this.noidung = noidung;
            this.tenuser = tenuser;
        }
    }
}
