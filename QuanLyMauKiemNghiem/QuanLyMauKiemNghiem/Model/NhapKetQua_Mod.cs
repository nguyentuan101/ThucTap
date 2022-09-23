using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyMauKiemNghiem.Object;

namespace QuanLyMauKiemNghiem.Model
{
    public class NhapKetQua_Mod
    {
        LoadData ld = new LoadData();
        public DataTable getMau()
        {
            string sql = "select distinct JOBNO,MAMAU,TENMAU,TIME_CHUYENMAU,TIME_TRABCPT_DK from PHIEU_YCPT where TRANGTHAI_CHUYEN=N'Đã chuyển'";
            return ld.GetData(sql);
        }
        public DataTable getChiTieuMau(string mamau)
        {
            string sql = "select MAPHIEU_YCPT,a.TENCT,a.TENNENMAU,a.MAPP,KETQUA_PT,GHICHU_PT,a.TENPP,a.MADV,a.TENDV,a.LOD,TRANGTHAI_KQ,NGAY_NHAP_KQ,b.PHAN_LOAI from PHIEU_YCPT a join CHI_TIEU b on a.MACT=b.MACT where MAMAU='" + mamau + "' and MAMAU<>''";
            
            return ld.GetData(sql);
        }
        public DataTable getPhuongPhap()
        {
            string sql = "select * from PHUONG_PHAP";
            return ld.GetData(sql);
        }
        public DataTable getTenPhuongPhap(string mapp)
        {
            string sql = "select TENPP from PHUONG_PHAP WHERE MAPP="+mapp;
            return ld.GetData(sql);
        }
        public DataTable getTenDonVi(string madv)
        {
            string sql = "select  TENDV from DON_VI where MADV=" + madv;
            return ld.GetData(sql);
        }
        public bool updateKetQua( Phieu_YCPT_Obj obj)
        {
            string sql = "update PHIEU_YCPT set MADV="+obj.Madv+",TENDV=N'"+obj.Tendv+"',LOD=N'"+obj.Lod+"', MAPP=" + obj.Mapp + ",TENPP=N'" + obj.Tenpp + "',KETQUA_PT=N'" + obj.Ketqua_pt + "',GHICHU_PT=N'" + obj.Ghichu_pt + "',TRANGTHAI_KQ=N'Chờ duyệt',NGAY_NHAP_KQ=GETDATE(),USER_TRA_KQ='" + obj.User_tra_kq + "' where MAPHIEU_YCPT=" + obj.Maphieu_ycpt;
            return ld.AddData(sql);
        }

        public DataTable getChiTieuKiemLai()
        {
            string sql = "select MAMAU,TENMAU,TENCT,TENNENMAU,LOD,TENDV,TENPP,GHICHU_DUYET,DUYET_KQPT,NGAY_DUYET from PHIEU_YCPT WHERE DUYET_KQPT=0";
            return ld.GetData(sql);
        }
        public bool thaydoiketqua(Phieu_YCPT_Obj obj)
        {
            string sql = "insert KET_QUA(MAPHIEU_YCPT,MAPP,TENPP,KETQUA_PT,GHICHU_PT,USER_TRA_KQ,NGAYGIO,MADV,TENDV,LOD) VALUES ("+obj.Maphieu_ycpt+","+obj.Mapp+",N'"+obj.Tenpp+"',N'"+obj.Ketqua_pt+"',N'"+obj.Ghichu_pt+"','"+obj.User_tra_kq+"',GETDATE(),"+obj.Madv+",N'"+obj.Tendv+"',N'"+obj.Lod+"');";
            return ld.AddData(sql);
        }
    }
}
