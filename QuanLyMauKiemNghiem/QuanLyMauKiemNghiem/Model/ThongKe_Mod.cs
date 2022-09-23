using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    public class ThongKe_Mod
    {
        LoadData ld = new LoadData();



        // THỐNG KÊ
        public DataTable GetData_ThongKeChiTieu(string tu, string den)
        {
            string sql = "SELECT TENCT,TENNENMAU, count(*) SOLANKIEM FROM PHIEU_YCPT where KETQUA_PT is not null and KETQUA_PT<>'' and (TIME_NHANMAU_DK between '" + tu + "' and '" + den + "') group by TENCT,TENNENMAU order by SOLANKIEM desc";
            return ld.GetData(sql);
        }

        public DataTable GetData_ThongKeChiTieuKH(string tu, string den, string makh)
        {
            string sql = "SELECT TENCT, count(*) SOLANKIEM FROM PHIEU_YCPT p join hop_dong hd on p.mahd=hd.mahd where KETQUA_PT is not null and KETQUA_PT<>'' and (TIME_NHANMAU_DK between '" + tu + "' and '" + den + "') and MAKH='" + makh + "' group by TENCT order by SOLANKIEM desc";
            return ld.GetData(sql);
        }

        public DataTable GetKH()
        {
            string sql = "select distinct kh.MAKH, kh.TENKH FROM KHACH_HANG kh join HOP_DONG hd on kh.MAKH=hd.MAKH join PHIEU_YCPT p on p.MAHD=hd.MAHD";
            return ld.GetData(sql);
        }
    }
}
