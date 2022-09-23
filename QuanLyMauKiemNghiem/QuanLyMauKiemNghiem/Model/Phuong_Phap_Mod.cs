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
   public class Phuong_Phap_Mod
    {
        LoadData ld = new LoadData();
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();
        public DataTable GetDataPP()
        {
            string sql = "Select * From PHUONG_PHAP";
            return ld.GetData(sql);
        }

        public bool addData(Phuong_Phap_Obj Obj)
        {
            string sql = "INSERT INTO dbo.PHUONG_PHAP(TENPP) VALUES  (N'" + Obj.Tenpp + "')";
            return ld.AddData(sql);
        }

        public bool update(Phuong_Phap_Obj Obj)
        {
            string sql = "UPDATE dbo.PHUONG_PHAP SET TENPP =N'" + Obj.Tenpp + "' where MAPP=N'" + Obj.Mapp + "'";
            return ld.AddData(sql);
        }

       // update bên bảng chỉ tiêu
        public bool update_CT(Phuong_Phap_Obj Obj)
        {
            string sql = "UPDATE dbo.CHI_TIEU SET TENPP =N'" + Obj.Tenpp + "' where MAPP=N'" + Obj.Mapp + "'";
            return ld.AddData(sql);
        }

        // update bên bảng phiếu yêu cầu phân tích
        public bool update_PYC(Phuong_Phap_Obj Obj)
        {
            string sql = "UPDATE dbo.PHIEU_YCPT SET TENPP =N'" + Obj.Tenpp + "' where MAPP=N'" + Obj.Mapp + "'";
            return ld.AddData(sql);
        }

        // update bên bảng công nợ
        public bool update_CN(Phuong_Phap_Obj Obj)
        {
            string sql = "UPDATE dbo.CONG_NO SET TENPP =N'" + Obj.Tenpp + "' where MAPP=N'" + Obj.Mapp + "'";
            return ld.AddData(sql);
        }

        // update bên bảng kết quả
        public bool update_KQ(Phuong_Phap_Obj Obj)
        {
            string sql = "UPDATE dbo.KET_QUA SET TENPP =N'" + Obj.Tenpp + "' where MAPP=N'" + Obj.Mapp + "'";
            return ld.AddData(sql);
        }

        public bool delectData(string ma)
        {
            string sql = "DELETE dbo.PHUONG_PHAP  where MAPP=N'" + ma + "'";
            return ld.AddData(sql);
        }

       // lấy ra mẫ pp lớn nhất
        public DataTable MaxMaPP()
        {
            string sql = "Select isnull(max(MADV), 0) as MADV  From DON_VI";
            return ld.GetData(sql);
        }
        // kiểm tra trùng trên đơn vị
        public DataTable kiemTraTenPP(string tennm)
        {
            string sql = "select * from PHUONG_PHAP where TENPP = N'" + tennm + "'";
            return ld.GetData(sql);
        }
        //KT trùng tên 
        public bool TrungTenPP(Phuong_Phap_Obj obj)
        {
            string sql = "select * from PHUONG_PHAP where TENPP = N'" + obj.Tenpp + "'";
            DataTable dt = ld.GetData(sql);
            int a = dt.Rows.Count;
            return (a > 0) ? true : false;
        }

        
    }
}
