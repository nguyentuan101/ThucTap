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
    public class Quy_Chuan_Mod
    {
        LoadData ld = new LoadData();
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();
        public DataTable GetDataQC()
        {
            string sql = "select * from QUY_CHUAN";
            return ld.GetData(sql);
        }

        public bool addData(Quy_Chuan_Obj Obj)
        {
            string sql = "INSERT INTO dbo.QUY_CHUAN(TENQC) VALUES  (N'" + Obj.Tenqc + "')";
            return ld.AddData(sql);
        }

        public bool update(Quy_Chuan_Obj Obj)
        {
            string sql = "UPDATE dbo.QUY_CHUAN SET TENQC =N'" + Obj.Tenqc + "' where MAQC=N'" + Obj.Maqc + "'";
            return ld.AddData(sql);
        }


        // update bên bảng chỉ tiêu
        public bool update_CT(Quy_Chuan_Obj Obj)
        {
            string sql = "UPDATE dbo.CHI_TIEU SET TENQC =N'" + Obj.Tenqc + "' where MAQC=N'" + Obj.Maqc + "'";
            return ld.AddData(sql);
        }

        // update bên bảng phiếu yêu cầu phân tích
        public bool update_PYC(Quy_Chuan_Obj Obj)
        {
            string sql = "UPDATE dbo.PHIEU_YCPT SET TENQC =N'" + Obj.Tenqc + "' where MAQC=N'" + Obj.Maqc + "'";
            return ld.AddData(sql);
        }

        // update bên bảng công nợ
        public bool update_CN(Quy_Chuan_Obj Obj)
        {
            string sql = "UPDATE dbo.CONG_NO SET TENQC =N'" + Obj.Tenqc + "' where MAQC=N'" + Obj.Maqc + "'";
            return ld.AddData(sql);
        }

        // update bên bảng kết quả
        public bool update_KQ(Quy_Chuan_Obj Obj)
        {
            string sql = "UPDATE dbo.KET_QUA SET TENQC =N'" + Obj.Tenqc + "' where MAQC=N'" + Obj.Maqc + "'";
            return ld.AddData(sql);
        }

        public bool delectData(string ma)
        {
            string sql = "DELETE dbo.QUY_CHUAN  where MAQC=" + ma;
            return ld.AddData(sql);
        }



        // kiểm tra trùng trên đơn vị
        public DataTable Check_TenQC(string tenqc)
        {
            string sql = "select * from QUY_CHUAN where TENQC = N'" + tenqc + "'";
            return ld.GetData(sql);
        }
        //KT trùng tên 
        public bool TrungTenQC(Quy_Chuan_Obj obj)
        {
            string sql = "select * from QUY_CHUAN where TENQC = N'" + obj.Tenqc + "'";
            DataTable dt = ld.GetData(sql);
            int a = dt.Rows.Count;
            return (a > 0) ? true : false;
        }


    }
}
