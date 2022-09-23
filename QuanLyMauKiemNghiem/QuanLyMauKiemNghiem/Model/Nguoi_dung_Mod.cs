using QuanLyMauKiemNghiem.Object;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace QuanLyMauKiemNghiem.Model
{
    public class Nguoi_dung_Mod
    {
        LoadData ld = new LoadData();
        //Phân quyền sử dụng
        public int phanquyen(string tenuser,int maquyen)
        {
            string sql = "select * from SD_QUYEN where TENUSER='"+tenuser+"' and IDQUYEN="+maquyen;
            return ld.GetData(sql).Rows.Count;
        }
        
        //
        public DataTable GetData()
        {
            string sql = "select * from NGUOI_DUNG where XOAND<>1 ";
            return ld.GetData(sql);
        }
        public DataTable GetQuyenSD(string tenuser)
        {
            string sql = "SELECT a.MAQUYEN,a.TENQUYEN, CASE WHEN a.MAQUYEN=b.IDQUYEN THEN 1 ELSE 0 END AS CHON FROM( [QLMKN2].[dbo].[DS_QUYEN] a LEFT JOIN (SELECT * FROM QLMKN2.dbo.SD_QUYEN WHERE TENUSER='"+tenuser+"') b ON a.MAQUYEN=b.IDQUYEN)";
            return ld.GetData(sql);
            
        }

        public DataTable Kiemtrataikhoan(string tenuser)
        {
            string sql = "select * from NGUOI_DUNG where TENUSER='"+tenuser+"'";
            return ld.GetData(sql);
        }
        public bool insertNguoiDung(Nguoi_Dung_Obj obj)
        {
            string sql = "insert NGUOI_DUNG(TENUSER,MATKHAU,HOTEN,NHOM,XOAND) values ('"+obj.Tenuser+"','"+obj.Matkhau+"',N'"+obj.Hoten+"',N'"+obj.Nhom+"',0)";
            return ld.AddData(sql);
        }
        public bool updateNguoiDung(Nguoi_Dung_Obj obj)
        {
            string sql = "update NGUOI_DUNG set MATKHAU='"+obj.Matkhau+"', HOTEN=N'"+obj.Hoten+"',NHOM=N'"+obj.Nhom+"' where TENUSER='"+obj.Tenuser+"'";
            return ld.AddData(sql);
        }
        public bool deleteNguoiDung(string tenuser)
        {
            string sql = "update NGUOI_DUNG set XOAND=1 WHERE TENUSER='"+tenuser+"'";
            return ld.AddData(sql);
        }
        public bool insertQuyenSD(SD_Quyen_Obj obj)
        {
            string sql = "insert SD_QUYEN(TENUSER,TENQUYEN,IDQUYEN) values ('"+obj.Tenuser+"',N'"+obj.Tenquyen+"',"+obj.Idquyen+")";
            return ld.AddData(sql);
        }
        public bool deleteQuyenSD(string tenuser,string idquyen)
        {
            string sql = "delete SD_QUYEN where TENUSER='"+tenuser+"' and IDQUYEN="+idquyen;
            return ld.AddData(sql);
        }
        public DataTable kiemtraQuyenSD(string tenuser, string idquyen)
        {
            string sql = "select * from SD_QUYEN where TENUSER='"+tenuser+"' and IDQUYEN="+idquyen;
            return ld.GetData(sql);
        }

        //Đổi mật khẩu
        public bool DoiMK(Nguoi_Dung_Obj Obj)
        {
            string sql = "UPDATE dbo.NGUOI_DUNG SET MATKHAU=N'" + Obj.Matkhau + "' WHERE TENUSER='" + Obj.Tenuser + "' and MATKHAU=N'" + Obj.Matkhaucu + "'";
            return ld.AddData(sql);
        }
    }
}
