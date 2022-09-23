using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    public class Nhan_Ban_JobNo_Mod
    {
        LoadData ld = new LoadData();

        public DataTable GetJOBNO()
        {
            string sql = "  select DISTINCT C.TENKH,A.MAHD,JOBNO from PHIEU_YCPT A JOIN HOP_DONG B ON A.MAHD=B.MAHD JOIN KHACH_HANG C ON B.MAKH=C.MAKH where JOBNO is not null AND JOBNO<>''";
            return ld.GetData(sql);
        }
        public DataTable GetMau(string jobno)
        {
            string sql = "select distinct MAMAU,TENMAU,TIME_NHANMAU_DK,MOTAMAU,NGUONMAU,KHOILUONGMAU,TINHTRANGMAU,NGUOINHANMAU,SEAL,TIME_TRABCPT_DK,TTKHCUNGCAP,temp_MAMAU from PHIEU_YCPT where JOBNO='" + jobno + "';";
            return ld.GetData(sql);
        }

        public DataTable GetChiTieu(string mamau)
        {
            string sql = "select MACT,TENCT,MACD,TENCD,MANENMAU,TENNENMAU,MADV,TENDV,LOD,MAPP,TENPP,DONGIA,NHOMCHITIEU,MAQC,TENQC from PHIEU_YCPT WHERE MAMAU='"+mamau+"'";
            return ld.GetData(sql);
        }
        public string KiemTraJobNo(string jobno,string mahd)
        {
            string sql = "exec Sp_KiemTraJobNo '"+jobno+"','"+mahd+"'";
            string temp = ld.GetData(sql).Rows[0][0].ToString();
            return temp;
        }
        public bool addData(Phieu_Nhan_Mau_Obj obj)
        {
            if (obj.Macd == "")
                obj.Macd = "NULL";
            if (obj.Manenmau == "")
                obj.Manenmau = "NULL";
            if (obj.Madv == "")
                obj.Madv = "NULL";
            if (obj.Mapp == "")
                obj.Mapp = "NULL";
            if (obj.Maqc == "")
                obj.Maqc = "NULL";
            string SQL;
            if (obj.Time_trabcpt_dk != "")
            {
                SQL = "insert into PHIEU_YCPT(MAHD,JOBNO,MAMAU,TENMAU,TIME_NHANMAU_DK,MOTAMAU,NGUONMAU,KHOILUONGMAU,TINHTRANGMAU,NGUOINHANMAU,SEAL,TIME_TRABCPT_DK,TTKHCUNGCAP,temp_MAMAU,MACT,TENCT,MACD,TENCD,MANENMAU,TENNENMAU,MADV,TENDV,LOD,MAPP,TENPP,DONGIA,NHOMCHITIEU,MAQC,TENQC,TIME_NHANMAU_TT, TRANGTHAI_CHUYEN) values ('" + obj.Mahd + "','" + obj.Jobno + "','" + obj.Mamau + "',N'" + obj.Tenmau + "','" + obj.Time_nhanmaudk + "',N'" + obj.Motamau + "',N'" + obj.Nguonmau + "',N'" + obj.Khoiluongmau + "',N'" + obj.Tinhtrangmau + "',N'" + obj.Nguoinhanmau + "','" + obj.Seal + "','" + obj.Time_trabcpt_dk + "',N'" + obj.Ttkhcungcap + "'," + obj.TempMaMau + ",'" + obj.Mact + "',N'" + obj.Tenct + "'," + obj.Macd + ",N'" + obj.Tencd + "'," + obj.Manenmau + ",N'" + obj.Tennenmau + "'," + obj.Madv + ",N'" + obj.Tendv + "','" + obj.Lod + "'," + obj.Mapp + ",N'" + obj.Tenpp + "'," + obj.Dongia + ",N'" + obj.Nhomct + "'," + obj.Maqc + ",N'" + obj.Tenqc + "',GETDATE(), N'Chưa chuyển')";
                if (obj.Jobno == "")
                    SQL = "insert into PHIEU_YCPT(MAHD,TENMAU,TIME_NHANMAU_DK,MOTAMAU,NGUONMAU,KHOILUONGMAU,TINHTRANGMAU,NGUOINHANMAU,SEAL,TIME_TRABCPT_DK,TTKHCUNGCAP,temp_MAMAU,MACT,TENCT,MACD,TENCD,MANENMAU,TENNENMAU,MADV,TENDV,LOD,MAPP,TENPP,DONGIA,NHOMCHITIEU,MAQC,TENQC,TRANGTHAI_CHUYEN) values ('" + obj.Mahd + "',N'" + obj.Tenmau + "','" + obj.Time_nhanmaudk + "',N'" + obj.Motamau + "',N'" + obj.Nguonmau + "',N'" + obj.Khoiluongmau + "',N'" + obj.Tinhtrangmau + "',N'" + obj.Nguoinhanmau + "','" + obj.Seal + "','" + obj.Time_trabcpt_dk + "',N'" + obj.Ttkhcungcap + "'," + obj.TempMaMau + ",'" + obj.Mact + "',N'" + obj.Tenct + "'," + obj.Macd + ",N'" + obj.Tencd + "'," + obj.Manenmau + ",N'" + obj.Tennenmau + "'," + obj.Madv + ",N'" + obj.Tendv + "','" + obj.Lod + "'," + obj.Mapp + ",N'" + obj.Tenpp + "'," + obj.Dongia + ",N'" + obj.Nhomct + "'," + obj.Maqc + ",N'" + obj.Tenqc + "', N'Chưa chuyển')";
            }
            else
            {
                SQL = "insert into PHIEU_YCPT(MAHD,JOBNO,MAMAU,TENMAU,TIME_NHANMAU_DK,MOTAMAU,NGUONMAU,KHOILUONGMAU,TINHTRANGMAU,NGUOINHANMAU,SEAL,TTKHCUNGCAP,temp_MAMAU,MACT,TENCT,MACD,TENCD,MANENMAU,TENNENMAU,MADV,TENDV,LOD,MAPP,TENPP,DONGIA,NHOMCHITIEU,MAQC,TENQC,TIME_NHANMAU_TT, TRANGTHAI_CHUYEN) values ('" + obj.Mahd + "','" + obj.Jobno + "','" + obj.Mamau + "',N'" + obj.Tenmau + "','" + obj.Time_nhanmaudk + "',N'" + obj.Motamau + "',N'" + obj.Nguonmau + "',N'" + obj.Khoiluongmau + "',N'" + obj.Tinhtrangmau + "',N'" + obj.Nguoinhanmau + "','" + obj.Seal + "',N'" + obj.Ttkhcungcap + "'," + obj.TempMaMau + ",'" + obj.Mact + "',N'" + obj.Tenct + "'," + obj.Macd + ",N'" + obj.Tencd + "'," + obj.Manenmau + ",N'" + obj.Tennenmau + "'," + obj.Madv + ",N'" + obj.Tendv + "','" + obj.Lod + "'," + obj.Mapp + ",N'" + obj.Tenpp + "'," + obj.Dongia + ",N'" + obj.Nhomct + "'," + obj.Maqc + ",N'" + obj.Tenqc + "',GETDATE(), N'Chưa chuyển')";
                if (obj.Jobno == "")
                    SQL = "insert into PHIEU_YCPT(MAHD,TENMAU,TIME_NHANMAU_DK,MOTAMAU,NGUONMAU,KHOILUONGMAU,TINHTRANGMAU,NGUOINHANMAU,SEAL,TTKHCUNGCAP,temp_MAMAU,MACT,TENCT,MACD,TENCD,MANENMAU,TENNENMAU,MADV,TENDV,LOD,MAPP,TENPP,DONGIA,NHOMCHITIEU,MAQC,TENQC, TRANGTHAI_CHUYEN) values ('" + obj.Mahd + "',N'" + obj.Tenmau + "','" + obj.Time_nhanmaudk + "',N'" + obj.Motamau + "',N'" + obj.Nguonmau + "',N'" + obj.Khoiluongmau + "',N'" + obj.Tinhtrangmau + "',N'" + obj.Nguoinhanmau + "','" + obj.Seal + "',N'" + obj.Ttkhcungcap + "'," + obj.TempMaMau + ",'" + obj.Mact + "',N'" + obj.Tenct + "'," + obj.Macd + ",N'" + obj.Tencd + "'," + obj.Manenmau + ",N'" + obj.Tennenmau + "'," + obj.Madv + ",N'" + obj.Tendv + "','" + obj.Lod + "'," + obj.Mapp + ",N'" + obj.Tenpp + "'," + obj.Dongia + ",N'" + obj.Nhomct + "'," + obj.Maqc + ",N'" + obj.Tenqc + "', N'Chưa chuyển')";
            }
            return ld.AddData(SQL);
        }

        public bool delMau(string tempMaMau)
        {
            return ld.AddData("delete PHIEU_YCPT where temp_MAMAU='" + tempMaMau + "'");
        }

        //
        //Nhân bản theo mẫu
        //
       


    }
}
