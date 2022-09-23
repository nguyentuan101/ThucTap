using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;


using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    public class Nhap_JobNo_Mod
    {
        LoadData ld = new LoadData();

        // LAY HOP DONG CHUA CO JOBNO
        public DataTable GetHopDong()
        {
            string sql = "SELECT diSTINCT HD.MAHD, TENHD, HD.MAKH, KH.TENKH, THOIHAN, NOIDUNG FROM HOP_DONG HD JOIN PHIEU_YCPT P  ON HD.MAHD=P.MAHD JOIN KHACH_HANG KH ON KH.MAKH=HD.MAKH WHERE MAMAU='' OR MAMAU IS NULL";
            return ld.GetData(sql);
        }


        // LAY MAU  CHUA CO JOBNO (với HĐ x)
        public DataTable GetMauChuaJobNo(string mahd)
        {
            string sql = "select distinct TENMAU,SLMAU,KHOILUONGMAU,TIME_NHANMAU_DK,temp_MAMAU from PHIEU_YCPT where mahd='" + mahd + "' and (JOBNO is null or JOBNO ='')";
            return ld.GetData(sql);
        }


        // LAY MAU ĐÃ CÓ JOBNO
        public DataTable GetMauCoJobNo()
        {
            string sql = " select distinct c.MAKH+'-'+c.TENKH as TENKH, b.MAHD +' '+TENHD as MAHD,JOBNO,MAMAU,TENMAU,SLMAU,KHOILUONGMAU,cast(TIME_NHANMAU_TT as date),TRANGTHAI_CHUYEN,temp_MAMAU from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on b.MAKH=c.MAKH where  JOBNO is not null and JOBNO<>'' ";
            return ld.GetData(sql);
        }


        // Tạo Mã Mẫu mới
        public int MaMau(string jobno)
        {
            string sql = "select temp_MAMAU from PHIEU_YCPT where JOBNO='" + jobno + "' group by temp_MAMAU";
            DataTable dt= ld.GetData(sql);
            int sl = dt.Rows.Count;
            return sl+1;
        }

        // Cap nhat JobNo
        public bool CapNhatJobNo(Phieu_YCPT_Obj phieu)
        {
            //GETDATE()
            string sql = "update PHIEU_YCPT set JOBNO='" + phieu.Jobno + "', MAMAU='" + phieu.Mamau + "', TIME_NHANMAU_TT='"+phieu.Time_nhanmau_tt+"' where temp_MAMAU=" + phieu.TempMaMau;
            return ld.AddData(sql);
        }

        //Kiem tra ton tại JobNo
        public string KiemTraJobNo(Phieu_YCPT_Obj phieu)
        {
            
             return ld.GetData("exec Sp_KiemTraJobNo '"+phieu.Jobno+"','"+phieu.Mahd+"'").Rows[0][0].ToString();
           
        }


        public bool Xoa_JobNo(string tempMamau)
        {
            string sql = "update PHIEU_YCPT set JOBNO=NULL, MAMAU=NULL where temp_MAMAU='"+tempMamau+"'";
            return ld.AddData(sql);
        }
        //LẤY TẤT CẢ CÁC MẪU THUỘC JOBNO
        public DataTable GetJobNo(string jobno)
        {
            string sql = "select distinct temp_MAMAU,a.MAMAU from PHIEU_YCPT a join (select distinct MAMAU from PHIEU_YCPT where JOBNO='"+jobno+"' group by MAMAU ) b on a.MAMAU=b.MAMAU";
            return ld.GetData(sql);
        }

        //Xuất phiếu yêu cầu phân tích
        public DataTable XuatPYCPT(string jobno)
        {
            string sql = "SELECT JOBNO, MAMAU,TENMAU,KHOILUONGMAU,MACT,TENCT,MAPP,TENPP FROM PHIEU_YCPT where JOBNO='" + jobno + "'";
            return ld.GetData(sql);
        }

        //Lấy nơi kiểm
        public DataTable getNoiKiem(string jobno)
        {
            string sql = "select distinct MANK,TENNK from PHIEU_YCPT where JOBNO='"+jobno+"'";
            return ld.GetData(sql);
        }

        public DataTable getChiTieuNhaThauPhu(string jobno,string nk)
        {
            string sql = "select JOBNO,MANK,TENNK,MAMAU,TENMAU,MACT,TENCT from PHIEU_YCPT where JOBNO='"+jobno+"' and MANK='"+nk+"'";
            return ld.GetData(sql);
        }

    }
}
