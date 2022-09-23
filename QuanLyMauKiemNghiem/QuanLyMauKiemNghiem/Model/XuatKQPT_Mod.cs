using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyMauKiemNghiem.Object;
using System.Data;
namespace QuanLyMauKiemNghiem.Model
{
    public class XuatKQPT_Mod
    {
        LoadData ld = new LoadData();
        public DataTable getJobNOChuaXuat()
        {
            string sql = "select distinct c.TENKH,a.MAHD,a.JOBNO from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on b.MAKH=c.MAKH where (TRANGTHAI_XUAT<>N'Đã xuất' OR TRANGTHAI_XUAT IS NULL) and TRANGTHAI_CHUYEN=N'Hoàn thành'";
            return ld.GetData(sql);
        }
        public DataTable getJobNODaXuat()
        {
            string sql = "select distinct TENKH,MAHD,JOBNO from CONG_NO";
            return ld.GetData(sql);
        }
        public DataTable getDSmauRAW()
        {
            string sql = "select distinct c.TENKH,b.MAHD,JOBNO,MAMAU,TENMAU from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on b.MAKH=c.MAKH where MAMAU is not null and MAMAU<>''";
            return ld.GetData(sql);
        }
        public DataTable getChiTieuRAW(string mamau)
        {
            string sql = "select JOBNO, c.TENKH,DIACHI,TTKHCUNGCAP,MOTAMAU,SEAL,NGUONMAU,TIME_NHANMAU_TT,TIME_CHUYENMAU,MAMAU,TENMAU,MACT,case when NHOMCHITIEU IS NULL or NHOMCHITIEU='' then TENCT else  NHOMCHITIEU end TENCT,TENCT AS TENCTPHULUC,MAPP,TENPP,MADV,TENDV,LOD,case when NHOMCHITIEU IS NULL or NHOMCHITIEU='' then KETQUA_PT else N'Xem phụ lục: '+MAMAU  end KETQUA_PT,KETQUA_PT AS KETQUA_PTPHULUC,MAQC ,TENQC,NHOMCHITIEU,MAPHIEU_YCPT,MACD,TENCD,MANENMAU,TENNENMAU,DONGIA,a.MAHD,SOLUONGKIEM,c.MAKH,TRANGTHAI_XUAT from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on c.MAKH=b.MAKH where MAMAU='" + mamau + "'";
            return ld.GetData(sql);
        }
        public DataTable getMau(string jobno)
        {
            string sql = "select distinct MAMAU,TENMAU from PHIEU_YCPT where JOBNO='" + jobno + "' and TRANGTHAI_CHUYEN=N'Hoàn thành'";
            return ld.GetData(sql);
        }
        
        public DataTable getMauDaXuat(string jobno)
        {
            string sql = "select distinct MAMAU,TENMAU,NGAYXUAT,LANKIEMTHU from CONG_NO where JOBNO='" + jobno + "'";
            return ld.GetData(sql);
        }
        public DataTable updateTrangThaiXuat(string mamau)
        {
            string sql = "update PHIEU_YCPT set TRANGTHAI_XUAT=N'Đã xuất',TRANGTHAI_CHUYEN=N'Đã xuất' where MAMAU='" + mamau + "'";
            return ld.GetData(sql);
        }

        public DataTable getThongTinMau(string mamau)
        {
            string sql = "select JOBNO, c.TENKH,DIACHI,TTKHCUNGCAP,MOTAMAU,SEAL,NGUONMAU,TIME_NHANMAU_TT,TIME_CHUYENMAU,MAMAU,TENMAU,MACT,case when NHOMCHITIEU IS NULL or NHOMCHITIEU='' then TENCT else  NHOMCHITIEU end TENCT,TENCT AS TENCTPHULUC,MAPP,TENPP,MADV,TENDV,LOD,case when NHOMCHITIEU IS NULL or NHOMCHITIEU='' then KETQUA_PT else N'Xem phụ lục: '+MAMAU  end KETQUA_PT,KETQUA_PT AS KETQUA_PTPHULUC,MAQC ,TENQC,NHOMCHITIEU,MAPHIEU_YCPT,MACD,TENCD,MANENMAU,TENNENMAU,DONGIA,a.MAHD,SOLUONGKIEM,c.MAKH,TRANGTHAI_XUAT,MANK,TENNK from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on b.MAKH=c.MAKH where MAMAU='" + mamau + "'";
            return ld.GetData(sql);
        }
        public DataTable getThongTinMauDaXuat(string mamau,string lankiem)
        {
            string sql = "select A.MAMAU,TENMAU,NHOMCHITIEU, JOBNO,B.TENKH,DIACHI,TTKHCUNGCAP,MOTAMAU,SEAL,NGUONMAU, TIME_NHANMAU_TT,TIME_CHUYENMAU,a.MACT,case when NHOMCHITIEU is null or NHOMCHITIEU='' then TENCT else NHOMCHITIEU end TENCT, TENCT AS TENCTPHULUC,MAPP,TENPP,MADV,TENDV,LOD, case when NHOMCHITIEU is null or NHOMCHITIEU='' then KETQUA_PT ELSE N'Xem phụ lục: '+A.MAMAU END KETQUA_PT ,KETQUA_PT AS KETQUAN_PTPHULUC,MAQC,TENQC,MAPHIEU_YCPT MACD,TENCD,MANENMAU,TENNENMAU,DONGIA,A.MAHD,SOLUONG FROM CONG_NO A JOIN KHACH_HANG B ON A.MAKH=B.MAKH JOIN ( SELECT distinct MAMAU, TTKHCUNGCAP,MOTAMAU,SEAL,NGUONMAU,TIME_CHUYENMAU FROM PHIEU_YCPT WHERE MAMAU='" + mamau + "') C ON A.MAMAU=C.MAMAU join(select MAMAU,MACT,MAX(LANKIEMTHU) lankiem from CONG_NO where LANKIEMTHU<=" + lankiem + " group by MAMAU,MACT) d on a.MAMAU=d.MAMAU and a.MACT=d.MACT and a.LANKIEMTHU=d.lankiem";
            return ld.GetData(sql);
        }
        //Mẫu đã xuất kết quả


    }
}
