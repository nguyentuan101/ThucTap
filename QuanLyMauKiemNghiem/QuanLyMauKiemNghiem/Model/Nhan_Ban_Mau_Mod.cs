using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyMauKiemNghiem.Object;
using System.Data;
namespace QuanLyMauKiemNghiem.Model
{
    public class Nhan_Ban_Mau_Mod
    {
        LoadData ld = new LoadData();
        public DataTable GetMau()
        {
            string sql = "select distinct temp_MAMAU,c.TENKH,b.MAHD,JOBNO,MAMAU,TENMAU,TIME_NHANMAU_DK from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on c.MAKH=b.MAKH";
            return ld.GetData(sql);
        }
        public DataTable LayThongTinMau(string tempMaMau)
        {
            string sql = "select distinct MAHD,JOBNO,MAMAU,TENMAU,KHOILUONGMAU,NGUONMAU,MOTAMAU,TIME_NHANMAU_DK,TIME_TRABCPT_DK,NGUOINHANMAU,TINHTRANGMAU,TTKHCUNGCAP,SEAL from PHIEU_YCPT where temp_MAMAU='"+tempMaMau+"'";
            return ld.GetData(sql);
        }
        public DataTable LayDSchiTieuMau(string tempMaMau)
        {
            string sql = "SELECT MACT,TENCT,MANENMAU,TENNENMAU,LOD,MADV,TENDV,MAPP,TENPP,MACD,TENCD,DONGIA,NHOMCHITIEU,MAQC,TENQC FROM PHIEU_YCPT WHERE temp_MAMAU='" + tempMaMau + "'";
            return ld.GetData(sql);
        }
    }
}
