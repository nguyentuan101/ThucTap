using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    public class HopDong_Mod
    {
        LoadData ld = new LoadData();

        // Select all HOP_DONG
        public DataTable GetHopDong()
        {
            string sql = "select MAHD, TENHD, a.MAKH, b.TENKH, THOIHAN, NOIDUNG, NGUOICHIUTN, TINHTRANG,MAHD as ID from hop_dong a join KHACH_HANG b on a.MAKH=b.MAKH";
            return ld.GetData(sql);
        }


        // Select all KHACH_HANG
        public DataTable GetKhachHang()
        {
            string sql = "select * from khach_hang";
            return ld.GetData(sql);
        }

        public bool UpdateHopDong(Hop_Dong_Obj Obj)
        {
            string sql = "update HOP_DONG set MAHD='"+Obj.Mahd+"', TENHD=N'"+Obj.Tenhd+"', MAKH='"+Obj.Makh+"', TENKH=N'"+Obj.Tenkh+"', THOIHAN=N'"+Obj.Thoihan+"', NOIDUNG=N'"+Obj.Noidung+"', NGUOICHIUTN=N'"+Obj.Nguoichiutn+"', TINHTRANG=N'"+Obj.Tinhtrang+"' where MAHD='"+Obj.Id+"'";
            return ld.AddData(sql);
        }


        // kiem tra ton tai HOP_DONG
        public int KTHopDong(Hop_Dong_Obj Obj)
        {
            string sql = "SELECT * FROM HOP_DONG WHERE MAHD='"+Obj.Mahd+"'";
            DataTable dt = ld.GetData(sql);
            int count = dt.Rows.Count;
            return count;
        }

        // THEM HOP DONG
        public bool ThemHopDong(Hop_Dong_Obj Obj)
        {
            string sql = "INSERT INTO HOP_DONG(MAHD, TENHD, MAKH,TENKH, THOIHAN, NOIDUNG, NGUOICHIUTN, TINHTRANG) VALUES ('" + Obj.Mahd + "',N'" + Obj.Tenhd + "','" + Obj.Makh + "',N'" + Obj.Tenkh + "',N'" + Obj.Thoihan + "',N'" + Obj.Noidung + "',N'" + Obj.Nguoichiutn + "',N'" + Obj .Tinhtrang+ "')";
            return ld.AddData(sql);
        }

        // Xóa
        public bool XoaHopDong(Hop_Dong_Obj Obj)
        {
            string sql = "";
            string sql1 = "select * from PHIEU_YCPT where MAHD='" + Obj.Mahd + "'";
            DataTable dt = ld.GetData(sql1);
            int i = dt.Rows.Count;
            if (i == 0)
            {
                 sql = "delete HOP_DONG where mahd='" + Obj.Mahd + "'";
                 return ld.AddData(sql);
            }
            return false;
            
        }
    }
}
