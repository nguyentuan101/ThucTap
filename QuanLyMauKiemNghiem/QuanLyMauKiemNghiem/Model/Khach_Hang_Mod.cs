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
    public class Khach_Hang_Mod
    {
        LoadData ld = new LoadData();
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();
        public DataTable GetDataKH()
        {
            string sql = "Select * From KHACH_HANG";
            return ld.GetData(sql);
        }

        public bool addData(Khach_Hang_Obj Obj)
        {
            string sql = "INSERT INTO dbo.KHACH_HANG( MAKH,TENKH ,DIACHI,SDT ,EMAIL,MASOTHUE ,SOFAX, NGUOI_LH, SDT_LH, EMAIL_LH) VALUES  ( N'" + Obj.Makh + "', N'" + Obj.Tenkh + "' , N'" + Obj.Diachi + "','" + Obj.Sdt + "' ,'" + Obj.Email + "','" + Obj.Masothue + "' , '" + Obj.Sofax + "', '"+Obj.Nguoi_lh+"', '"+Obj.Sdt_lh+"', '"+Obj.Email_lh+"'  )";
            return ld.AddData(sql);
        }

        public bool update(Khach_Hang_Obj Obj)
        {
            string sql = "UPDATE dbo.KHACH_HANG SET TENKH =N'" + Obj.Tenkh + "' , DIACHI=N'" + Obj.Diachi + "',MASOTHUE='" + Obj.Masothue + "',SDT='" + Obj.Sdt + "' ,EMAIL='" + Obj.Email + "' , SOFAX ='" + Obj.Sofax + "', NGUOI_LH = '" + Obj.Nguoi_lh + "', SDT_LH = '" + Obj.Sdt_lh + "', EMAIL_LH = '" + Obj.Email_lh + "' where MAKH=N'" + Obj.Makh + "'";
            return ld.AddData(sql);
        }

        public bool delectData(string ma)
        {
            string sql = "DELETE dbo.KHACH_HANG  where MAKH=N'" + ma + "'";
            return ld.AddData(sql);
        }

        // kiểm tra xem khách hàng có trong bảng hợp đồng hay ko
        public DataTable check_Xoa_kh(string makh)
        {
            string sql = "select count(MAKH) as MAKH from HOP_DONG where MAKH = '"+makh+"' ";
            return ld.GetData(sql);
        }

    }
}
