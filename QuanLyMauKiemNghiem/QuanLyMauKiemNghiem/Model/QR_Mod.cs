using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
    public class QR_Mod
    {
        LoadData ld = new LoadData();
        public DataTable KiemTramau(string mamau,bool kiemtraxuat)
        {
            string sql = "select top 1 MAMAU +' '+TENMAU ten,TRANGTHAI_CHUYEN,MACT from PHIEU_YCPT where MAMAU='" + mamau + "'";   
            if(kiemtraxuat==true)
                sql = "select MAMAU +' '+TENMAU ten,TRANGTHAI_CHUYEN,MACT from PHIEU_YCPT where TRANGTHAI_CHUYEN=N'Kiểm lại' and MAMAU='" + mamau + "'";   
            return ld.GetData(sql);
        }

        public bool updateTrangThaiChuyen(string mamau)
        {
            string sql = "update PHIEU_YCPT set TRANGTHAI_KQ=N'Chưa kiểm' where MAMAU='" + mamau + "' and (TRANGTHAI_KQ is null or TRANGTHAI_KQ=''); update PHIEU_YCPT set TRANGTHAI_CHUYEN=N'Đã chuyển', TIME_CHUYENMAU=GETDATE() where MAMAU='" + mamau + "'";
            return ld.AddData(sql);
        }

        public DataTable getDSMauCoJobNo()
        {
            string sql = "SELECT DISTINCT MAHD,JOBNO,A.MAMAU,TENMAU,TRANGTHAI_CHUYEN,CONCAT(A.MAMAU,' ',TENMAU,N', Ngày nhận mẫu: ',FORMAT(TIME_NHANMAU_TT,'dd/MM/yyyy'),N', Số chỉ tiêu: ',SOCT,N' ,Ngày trả kết quả dự kiến: ', FORMAT(TIME_TRABCPT_DK,'dd/MM/yyyy')) as QR FROM PHIEU_YCPT A JOIN (select MAMAU,COUNT(*) SOCT from PHIEU_YCPT  WHERE MAMAU IS NOT NULL and MAMAU<>'' group by MAMAU) B ON A.MAMAU=B.MAMAU";

            return ld.GetData(sql);
        }
    }
}
