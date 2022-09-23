using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyMauKiemNghiem.Object;

namespace QuanLyMauKiemNghiem.Model
{
    public class Chi_Dinh_Mod
    {
        LoadData ld = new LoadData();

        public DataTable GetDataChiDinh(string macd)
        {
            string sql = "Select * From CHI_DINH";
            if (macd != null)
                sql = "select * from PHIEU_YCPT where MACD="+macd;
            return ld.GetData(sql);
        }
        public DataTable getDataChiDinhChiTieu(string macd)
        {
            string sql = "select MACD,A.MACT,TENCD,A.TENCT,MANENMAU,TENNENMAU,LOD,MADV,TENDV from CHI_DINH_CHI_TIEU a join CHI_TIEU b on a.MACT=b.MACT where MACD='" + macd + "'";
            return ld.GetData(sql);
        }
        public DataTable getDataChiTieu(string macd)
        {
            string sql = "SELECT MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV from CHI_TIEU where MACT NOT IN (SELECT MACT FROM CHI_DINH_CHI_TIEU WHERE MACD='"+macd+"')";
            return ld.GetData(sql);
        }

        public DataTable GetDataNenMau()
        {
            string sql = "select distinct MANENMAU,TENNENMAU from CHI_TIEU";
            return ld.GetData(sql);
        }
        public DataTable GetDataChiTieuNenMau()
        {
            string sql = "select * from CHI_TIEU ";
            return ld.GetData(sql);
        }
        public DataTable KiemTraTenCHiDinh(string tenchidinh,string macd)
        {
            string sql = "select * from CHI_DINH where TENCD=N'" + tenchidinh + "' and MACD<>" + macd;
            if(macd==null)
             sql = "select MACD,TENCD from CHI_DINH where TENCD=N'" + tenchidinh + "'";
            return ld.GetData(sql);
        }

        public bool addDataChiDinh(string tencd)
        {
            string sql = "insert into CHI_DINH (TENCD) values (N'"+tencd+"')";
            return ld.AddData(sql);
        }

        public bool addDataChiDinhCT(Chi_Dinh_Chi_Tieu_Obj Obj)
        {
            string sql = "insert into CHI_DINH_CHI_TIEU(MACD,TENCD,MACT,TENCT) values (" + Obj.Macd + ",N'" + Obj.Tencd + "','" + Obj.Mact + "',N'" + Obj.Tenct + "')";
            return ld.AddData(sql);
        }
        public bool delectDataCTCD(Chi_Dinh_Chi_Tieu_Obj Obj)
        {
            string sql = "delete CHI_DINH_CHI_TIEU where MACD='"+Obj.Macd+"' and MACT='"+Obj.Mact+"'";
            return ld.AddData(sql);
        }

        public bool update(Chi_Dinh_Obj Obj)
        {
            string sql = "update CHI_DINH set TENCD=N'" + Obj.Tencd + "' where MACD=" + Obj.Macd + "; update CHI_DINH_CHI_TIEU set TENCD=N'" + Obj.Tencd + "' where MACD=" + Obj.Macd + "; update PHIEU_YCPT set TENCD=N'" + Obj.Tencd + "' where MACD=" + Obj.Macd + "; update KET_QUA set TENCD=N'" + Obj.Tencd + "' where MACD=" + Obj.Macd + "; update CONG_NO set TENCD=N'" + Obj.Tencd + "' where MACD=" + Obj.Macd + ";";
            return ld.AddData(sql);
        }
        public bool delectCD(string macd)
        {
            string sql = "delete CHI_DINH where MACD="+macd;
            return ld.AddData(sql);
        }
        //KT trùng tên 
        public bool TrungTenCD(Chi_Dinh_Obj obj)
        {
            string sql = "select * from CHI_DINH where TENCD = N'" + obj.Tencd + "'";
            DataTable dt = ld.GetData(sql);
            int a = dt.Rows.Count;
            return (a > 0) ? true : false;
        }
        public bool addData(Chi_Dinh_Obj obj)
        {
            string sql = "insert into CHI_DINH (TENCD) values (N'" + obj.Tencd + "')";
            return ld.AddData(sql);
        }

    }
}
