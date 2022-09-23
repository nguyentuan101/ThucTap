using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    class PhanLoaiGia_Mod
    {
        LoadData ld = new LoadData();

        public DataTable GetPhanLoaiGia()
        {
            string sql = "select * from phan_loai_gia";
            return ld.GetData(sql);
        }

        public bool ThemPhanLoaiGia(Phan_Loai_Gia_Obj Obj)
        {
            string sql = "insert into phan_loai_gia(tenphanloaigia) values(N'" + Obj.Tenphanloaigia + "')";
            return ld.AddData(sql);
        }

        public bool SuaPhanLoaiGia(Phan_Loai_Gia_Obj Obj)
        {
            string sql = "update phan_loai_gia set tenphanloaigia=N'" + Obj.Tenphanloaigia + "' where maphanloaigia='" + Obj.Maphanloaigia + "'";
            return ld.AddData(sql);
        }


        public bool XoaPhanLoaiGia(Phan_Loai_Gia_Obj Obj)
        {
            string sql = "delete from phan_loai_gia where maphanloaigia='"+Obj.Maphanloaigia+"'";
            return ld.AddData(sql);
        }


        public int KTPhanLoaiGia(Phan_Loai_Gia_Obj Obj)
        {
            string sql = "SELECT * FROM PHAN_LOAI_GIA WHERE TENPHANLOAIGIA LIKE N'" + Obj.Tenphanloaigia + "'";
            DataTable dt = ld.GetData(sql);
            int count = dt.Rows.Count;
            return count;
        }
    }
}
