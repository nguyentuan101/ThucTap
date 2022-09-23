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
    public class Nen_Mau_Mod
    {
        LoadData ld = new LoadData();
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();
        public DataTable GetDataNM()
        {
            string sql = "select * from NEN_MAU";
            return ld.GetData(sql);
        }

        public bool addData(Nen_Mau_Obj Obj)
        {
            string sql = "INSERT INTO dbo.NEN_MAU(TENNENMAU) VALUES  (N'" + Obj.Tennenmau + "')";
            return ld.AddData(sql);
        }

        public bool update(Nen_Mau_Obj Obj)
        {
            string sql = "UPDATE dbo.NEN_MAU SET TENNENMAU =N'" + Obj.Tennenmau + "' where MANENMAU=N'" + Obj.Manenmau + "'";
            return ld.AddData(sql);
        }

        // update bên bảng chỉ tiêu
        public bool update_CT(Nen_Mau_Obj Obj)
        {
            string sql = "UPDATE dbo.CHI_TIEU SET TENNENMAU =N'" + Obj.Tennenmau + "' where MANENMAU=N'" + Obj.Manenmau + "'";
            return ld.AddData(sql);
        }

        // update bên bảng phiếu yêu cầu phân tích
        public bool update_PYC(Nen_Mau_Obj Obj)
        {
            string sql = "UPDATE dbo.PHIEU_YCPT SET TENNENMAU =N'" + Obj.Tennenmau + "' where MANENMAU=N'" + Obj.Manenmau + "'";
            return ld.AddData(sql);
        }

        // update bên bảng công nợ
        public bool update_CN(Nen_Mau_Obj Obj)
        {
            string sql = "UPDATE dbo.CONG_NO SET TENNENMAU =N'" + Obj.Tennenmau + "' where MANENMAU=N'" + Obj.Manenmau + "'";
            return ld.AddData(sql);
        }

        // update bên bảng kết quả
        public bool update_KQ(Nen_Mau_Obj Obj)
        {
            string sql = "UPDATE dbo.KET_QUA SET TENNENMAU =N'" + Obj.Tennenmau + "' where MANENMAU=N'" + Obj.Manenmau + "'";
            return ld.AddData(sql);
        }

        public bool delectData(string ma)
        {
            string sql = "DELETE dbo.NEN_MAU  where MANENMAU=N'" + ma + "'";
            return ld.AddData(sql);
        }



        // kiểm tra trùng trên đơn vị
        public DataTable Check_TenNM(string tennm)
        {
            string sql = "select * from NEN_MAU where TENNENMAU = N'" + tennm + "'";
            return ld.GetData(sql);
        }
        //KT trùng tên 
        public bool TrungTenNM(Nen_Mau_Obj obj)
        {
            string sql = "select * from NEN_MAU where TENNENMAU = N'" + obj.Tennenmau + "'";
            DataTable dt = ld.GetData(sql);
            int a = dt.Rows.Count;
            return (a > 0) ? true : false;
        }
    }
}
