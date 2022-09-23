using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    class Lookup_Ket_Qua_Obj
    {
        private int id, idphieunhan;

        public int Idphieunhan
        {
            get { return idphieunhan; }
            set { idphieunhan = value; }
        }

        public int Id
        {
            get { return id; }
            set { id = value; }
        }
        private string mahd, jobno, mamau, tenmau, chitieu, nenmau, phuongphap, ketquacu, ketquamoi, ngay, tenuser, donvi, phanloai, lod, ghichukq;

        public string Ghichukq
        {
            get { return ghichukq; }
            set { ghichukq = value; }
        }

        public string Lod
        {
            get { return lod; }
            set { lod = value; }
        }

        public string Phanloai
        {
            get { return phanloai; }
            set { phanloai = value; }
        }

        public string Donvi
        {
            get { return donvi; }
            set { donvi = value; }
        }

        public string Tenuser
        {
            get { return tenuser; }
            set { tenuser = value; }
        }

        public string Ngay
        {
            get { return ngay; }
            set { ngay = value; }
        }

        public string Ketquamoi
        {
            get { return ketquamoi; }
            set { ketquamoi = value; }
        }

        public string Ketquacu
        {
            get { return ketquacu; }
            set { ketquacu = value; }
        }

        public string Phuongphap
        {
            get { return phuongphap; }
            set { phuongphap = value; }
        }

        public string Nenmau
        {
            get { return nenmau; }
            set { nenmau = value; }
        }

        public string Chitieu
        {
            get { return chitieu; }
            set { chitieu = value; }
        }

        public string Tenmau
        {
            get { return tenmau; }
            set { tenmau = value; }
        }

        public string Mamau
        {
            get { return mamau; }
            set { mamau = value; }
        }

        public string Jobno
        {
            get { return jobno; }
            set { jobno = value; }
        }

        public string Mahd
        {
            get { return mahd; }
            set { mahd = value; }
        }

        public Lookup_Ket_Qua_Obj()
        {
        }
    }
}
