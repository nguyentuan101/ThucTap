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
    public class Don_Vi_Mod
    {
        LoadData ld = new LoadData();
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();
        public DataTable GetDataDV()
        {
            string sql = "Select * From DON_VI";
            return ld.GetData(sql);
        }

        public bool addData(Don_Vi_Obj Obj)
        {
            string sql = "INSERT INTO dbo.DON_VI(TENDV) VALUES  (N'" + Obj.Tendv + "')";
            return ld.AddData(sql);
        }

        public bool update(Don_Vi_Obj Obj)
        {
            string sql = "UPDATE dbo.DON_VI SET TENDV =N'" + Obj.Tendv + "' where MADV=N'" + Obj.Madv + "'";
            return ld.AddData(sql);
        }

        // update bên bảng chỉ tiêu
        public bool update_CT(Don_Vi_Obj Obj)
        {
            string sql = "UPDATE dbo.CHI_TIEU SET TENDV =N'" + Obj.Tendv + "' where MADV=N'" + Obj.Madv + "'";
            return ld.AddData(sql);
        }

        // update bên bảng phiếu yêu cầu phân tích
        public bool update_PYC(Don_Vi_Obj Obj)
        {
            string sql = "UPDATE dbo.PHIEU_YCPT SET TENDV =N'" + Obj.Tendv + "' where MADV=N'" + Obj.Madv + "'";
            return ld.AddData(sql);
        }

        // update bên bảng công nợ
        public bool update_CN(Don_Vi_Obj Obj)
        {
            string sql = "UPDATE dbo.CONG_NO SET TENDV =N'" + Obj.Tendv + "' where MADV=N'" + Obj.Madv + "'";
            return ld.AddData(sql);
        }

        // update bên bảng kết quả
        public bool update_KQ(Don_Vi_Obj Obj)
        {
            string sql = "UPDATE dbo.KET_QUA SET TENDV =N'" + Obj.Tendv + "' where MADV=N'" + Obj.Madv + "'";
            return ld.AddData(sql);
        }

        public bool delectData(string ma)
        {
            string sql = "DELETE dbo.DON_VI  where MADV=N'" + ma + "'";
            return ld.AddData(sql);
        }



        // kiểm tra trùng trên đơn vị
        public DataTable Check_TenDV(string tendv)
        {
            string sql = "select * from DON_VI where TENDV = N'"+tendv+"'";
            return ld.GetData(sql);
        }

   
        //KT trùng tên 
        public bool TrungTenDV(Don_Vi_Obj obj)
        {
            string sql = "select * from DON_VI where TENDV = N'" + obj.Tendv + "'";
            DataTable dt = ld.GetData(sql);
            int a = dt.Rows.Count;
            return (a > 0) ? true : false;
        }
    }
}
