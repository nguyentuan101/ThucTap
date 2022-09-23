using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    public class DuyetKQ_Mod
    {
        LoadData ld = new LoadData();


        public DataTable GetKetQua()
        {
            string sql = "select MAPHIEU_YCPT, MACT, TENCT, MANENMAU, TENNENMAU, MAPP,TENPP, KETQUA_PT,TENDV,LOD,GHICHU_PT,MAMAU+' - '+TENMAU AS MAU,MAMAU FROM PHIEU_YCPT WHERE TRANGTHAI_KQ=N'Chờ duyệt' and TRANGTHAI_CHUYEN=N'Đã chuyển'";
            return ld.GetData(sql);
        }


        public bool Duyet_Dat(Phieu_YCPT_Obj phieu)
        {
            string sql = "update PHIEU_YCPT set DUYET_KQPT='1', GHICHU_DUYET=N'" + phieu.Ghichu_duyet + "', NGAY_DUYET=GETDATE(), USER_DUYET='" + phieu.User_duyet + "',TRANGTHAI_KQ=N'Hoàn thành' where MAPHIEU_YCPT='" + phieu.Maphieu_ycpt + "'";
            return ld.AddData(sql);
        }


        public bool Duyet_Khong_Dat(Phieu_YCPT_Obj phieu)
        {
            string sql = "update PHIEU_YCPT set DUYET_KQPT='0', GHICHU_DUYET=N'" + phieu.Ghichu_duyet + "', NGAY_DUYET='" + phieu.Ngay_duyet + "', USER_DUYET='" + phieu.User_duyet + "',TRANGTHAI_KQ=N'Kiểm lại' where MAPHIEU_YCPT='" + phieu.Maphieu_ycpt + "'";
            return ld.AddData(sql);
        }
        public DataTable KiemTraHoanThanh(string mamau)
        {
            string sql = "select MAPHIEU_YCPT from PHIEU_YCPT where MAMAU='" + mamau + "' and MAPHIEU_YCPT not in (select MAPHIEU_YCPT from PHIEU_YCPT where MAMAU='"+mamau+"' and TRANGTHAI_KQ=N'Hoàn thành')";
            return ld.GetData(sql);
        }
        public bool UpdateTrangThaiChuyen(string mamau,bool nhanmau)
        {   string A = DateTime.Now.ToString("yyyy-MM-dd HH:MM");
            string sql = "update PHIEU_YCPT set TRANGTHAI_CHUYEN=N'Hoàn thành' ,TIME_TRABCPT_TT='"+A+"' where MAMAU='" + mamau + "'";
            if(nhanmau==true)
                sql = "update PHIEU_YCPT set TRANGTHAI_CHUYEN=N'Chưa chuyển' where MAMAU='" + mamau + "'";
            return ld.AddData(sql);
        }

        //DANH SÁCH THEO DÕI KIỂM NGHIỆM
        public DataTable dsTheoDoiKiemNghiem()
        {
            string sql = " select JOBNO,MAMAU,TIME_TRABCPT_DK,TIME_TRABCPT_TT,TENNENMAU,TENCT,TENPP,TENNK,TENDV,LOD,KETQUA_PT,TRANGTHAI_KQ from PHIEU_YCPT where TRANGTHAI_KQ is not null or TRANGTHAI_KQ<>''";
            return ld.GetData(sql);
        }

        public bool thaydoiketquaduyet(Phieu_YCPT_Obj obj,bool duyet)
        {
            string sql = "insert KET_QUA(MAPHIEU_YCPT,MAPP,TENPP,KETQUA_PT,GHICHU_PT,DUYET_KQPT,GHICHU_DUYET,NGAY_DUYET,USER_DUYET) VALUES ("+obj.Maphieu_ycpt+","+obj.Mapp+",N'"+obj.Tenpp+"',N'"+obj.Ketqua_pt+"',N'"+obj.Ghichu_pt+"',1,'"+obj.Ghichu_duyet+"',GETDATE(),'"+obj.User_duyet+"')";
            if(duyet==false)
                sql = "insert KET_QUA(MAPHIEU_YCPT,MAPP,TENPP,KETQUA_PT,GHICHU_PT,DUYET_KQPT,GHICHU_DUYET,NGAY_DUYET,USER_DUYET) VALUES (" + obj.Maphieu_ycpt + "," + obj.Mapp + ",N'" + obj.Tenpp + "',N'" + obj.Ketqua_pt + "',N'" + obj.Ghichu_pt + "',0,'" + obj.Ghichu_duyet + "',GETDATE(),'" + obj.User_duyet + "')";
            return ld.AddData(sql);
        }
       
    }
}
