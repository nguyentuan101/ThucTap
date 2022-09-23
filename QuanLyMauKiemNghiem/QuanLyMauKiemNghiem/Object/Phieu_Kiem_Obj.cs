using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Phieu_Kiem_Obj
    {
        string idphieunhan, idchitieu, idchidinh, ngaytra_kqpt, nd_thamdinh, user_duyet, ngayduyet, ngaychuyenmau, user_tra_kq, ghichukq;

        public string Idphieunhan
        {
            get { return idphieunhan; }
            set { idphieunhan = value; }
        }

        public string Idchitieu
        {
            get { return idchitieu; }
            set { idchitieu = value; }
        }

        public string Idchidinh
        {
            get { return idchidinh; }
            set { idchidinh = value; }
        }

        public string Ngaytra_kqpt
        {
            get { return ngaytra_kqpt; }
            set { ngaytra_kqpt = value; }
        }

        public string Nd_thamdinh
        {
            get { return nd_thamdinh; }
            set { nd_thamdinh = value; }
        }

        public string User_duyet
        {
            get { return user_duyet; }
            set { user_duyet = value; }
        }

        public string Ngayduyet
        {
            get { return ngayduyet; }
            set { ngayduyet = value; }
        }

        public string Ngaychuyenmau
        {
            get { return ngaychuyenmau; }
            set { ngaychuyenmau = value; }
        }

        public string User_tra_kq
        {
            get { return user_tra_kq; }
            set { user_tra_kq = value; }
        }

        public string Ghichukq
        {
            get { return ghichukq; }
            set { ghichukq = value; }
        }
        int soluong, idnoikiem, idpp;

        public int Soluong
        {
            get { return soluong; }
            set { soluong = value; }
        }

        public int Idnoikiem
        {
            get { return idnoikiem; }
            set { idnoikiem = value; }
        }

        public int Idpp
        {
            get { return idpp; }
            set { idpp = value; }
        }
        float dongia;

        public float Dongia
        {
            get { return dongia; }
            set { dongia = value; }
        }
        bool thamdinh, xuat_kq;

        public bool Thamdinh
        {
            get { return thamdinh; }
            set { thamdinh = value; }
        }

        public bool Xuat_kq
        {
            get { return xuat_kq; }
            set { xuat_kq = value; }
        }


        public Phieu_Kiem_Obj()
        {

        }

    }
}
