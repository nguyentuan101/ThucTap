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
    public class Noi_Kiem_Mod
    {
        LoadData ld = new LoadData();
        ConnectToSQL con = new ConnectToSQL();
        SqlCommand cmd = new SqlCommand();
        public DataTable GetDataNK()
        {
            string sql = "select * from NOI_KIEM";
            return ld.GetData(sql);
        }

        public bool addData(Noi_Kiem_Obj Obj)
        {
            string sql = "INSERT INTO dbo.NOI_KIEM(TENNK) VALUES  (N'" + Obj.Tennoikiem + "')";
            return ld.AddData(sql);
        }

        public bool update(Noi_Kiem_Obj Obj)
        {
            string sql = "UPDATE dbo.NOI_KIEM SET TENNK =N'" + Obj.Tennoikiem + "' where MANK=N'" + Obj.Idkiem + "'";
            return ld.AddData(sql);
        }


        // update bên bảng chỉ tiêu
        public bool update_CT(Noi_Kiem_Obj Obj)
        {
            string sql = "UPDATE dbo.CHI_TIEU SET TENNK =N'" + Obj.Tennoikiem + "' where MANK=N'" + Obj.Idkiem + "'";
            return ld.AddData(sql);
        }

        // update bên bảng phiếu yêu cầu phân tích
        public bool update_PYC(Noi_Kiem_Obj Obj)
        {
            string sql = "UPDATE dbo.PHIEU_YCPT SET TENNK =N'" + Obj.Tennoikiem + "' where MANK=N'" + Obj.Idkiem + "'";
            return ld.AddData(sql);
        }

        // update bên bảng công nợ
        public bool update_CN(Noi_Kiem_Obj Obj)
        {
            string sql = "UPDATE dbo.CONG_NO SET TENNK =N'" + Obj.Tennoikiem + "' where MANK=N'" + Obj.Idkiem + "'";
            return ld.AddData(sql);
        }

        // update bên bảng kết quả
        public bool update_KQ(Noi_Kiem_Obj Obj)
        {
            string sql = "UPDATE dbo.KET_QUA SET TENNK =N'" + Obj.Tennoikiem + "' where MANK=N'" + Obj.Idkiem + "'";
            return ld.AddData(sql);
        }

        public bool delectData(string ma)
        {
            string sql = "DELETE dbo.NOI_KIEM  where MANK=N'" + ma + "'";
            return ld.AddData(sql);
        }



        // kiểm tra trùng trên đơn vị
        public DataTable Check_TenNK(string tennk)
        {
            string sql = "select * from NOI_KIEM where TENNK = N'" + tennk + "'";
            return ld.GetData(sql);
        }
    }
}
