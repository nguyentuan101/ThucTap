using QuanLyMauKiemNghiem.Object;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
   public class Nhan_Mau_Mod
    {
       LoadData ld = new LoadData();

       // lấy danh sách hợp đồng
       public DataTable GetHopDong()
       {
           string sql = "select MAHD, kh.TENKH,TENHD from HOP_DONG as hd, KHACH_HANG as kh where hd.MAKH = kh.MAKH";
           return ld.GetData(sql);
       }
       public string GetTenKH(string mahd)
       {
           string sql = "select b.TENKH from HOP_DONG a join KHACH_HANG b on a.MAKH=b.MAKH where MAHD='"+mahd+"'";
           return ld.GetData(sql).Rows[0][0].ToString();
       }
       // lấy id của mẫu khi xác nhận
       public DataTable GetIDMau()
       {
           string sql = "select MAX(MAPHIEU_YCPT) as MA from PHIEU_YCPT";
           return ld.GetData(sql);
       }

       // lấy danh sách chỉ định
       public DataTable GetChiDinh()
       {
           string sql = "select distinct MACD, TENCD from CHI_DINH_CHI_TIEU as ctcd, CHI_TIEU as ct where ctcd.MACT = ct.MACT";
           return ld.GetData(sql);
       }

       // lấy danh sách chỉ tiêu
       public DataTable GetChiTieu(string mahd, string tenmau, string ngay)
       {
           string sql = "select ct.MACT, ct.TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, gct.GIA, MAPP, TENPP from CHI_TIEU as ct, GIA_CHI_TIEU as gct where ct.MACT NOT IN (select ISNULL(MACT,'') from PHIEU_YCPT where MAHD = N'"+mahd+"' and TENMAU = N'"+tenmau+"' and TIME_NHANMAU_DK = N'"+ngay+"')";
           return ld.GetData(sql);
       }
       // lấy danh sách chỉ tiêu 2
       public DataTable GetChiTieu2(string tempMaMau)
       {
           string sql = "SELECT a.MACT,a.TENCT,MANENMAU, TENNENMAU,LOD,MADV,TENDV,GIA,MAPP,TENPP,NHOMCHITIEU,TENPHANLOAIGIA,MAQC,TENQC FROM CHI_TIEU a left join GIA_CHI_TIEU b on a.MACT=b.MACT where a.MACT not in (select MACT from PHIEU_YCPT where temp_MAMAU=" + tempMaMau + " and MACT is not null)";
           return ld.GetData(sql);
       }
      
       //Kiểm tra chỉ tieu đã có trong mẫu chưa
       public DataTable KiemTraChiTieu(Phieu_Nhan_Mau_Obj obj)
       {
           string sql = "select * from PHIEU_YCPT where  temp_MAMAU=" + obj.TempMaMau + " and MACT='" + obj.Mact + "'";
           return ld.GetData(sql);
       }
       // lấy danh sách chỉ tiêu theo chỉ định
       public DataTable GetChiTieu_ChiDinh(string macd, Phieu_Nhan_Mau_Obj obj)
       {
           string sql = "select ctcd.MACT, ctcd.TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, MACD, TENCD, MAPP, TENPP ,NHOMCHITIEU, TENPHANLOAIGIA, MAQC, TENQC, GIA from CHI_DINH_CHI_TIEU as ctcd, CHI_TIEU as ct left join GIA_CHI_TIEU b on ct.MACT=b.MACT where ctcd.MACT = ct.MACT and MACD = '"+macd+"' and ct.MACT not in (select MACT from PHIEU_YCPT where temp_MAMAU= '"+obj.TempMaMau+"' and MACT is not null)";
           return ld.GetData(sql);
       }

       public bool addData(Phieu_Nhan_Mau_Obj Obj)
       {
           string sql = "insert into PHIEU_YCPT(MAHD, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, TTKHCUNGCAP, NGUOINHANMAU, SEAL, TIME_NHANMAU_DK,TIME_TRABCPT_DK,temp_MAMAU,JOBNO,MAMAU) VALUES ( N'" + Obj.Mahd + "', N'" + Obj.Tenmau + "', N'" + Obj.Motamau + "', N'" + Obj.Nguonmau + "', N'" + Obj.Khoiluongmau + "', N'" + Obj.Tinhtrangmau + "', N'" + Obj.Ttkhcungcap + "',N'" + Obj.Nguoinhanmau + "', N'" + Obj.Seal + "', N'" + Obj.Time_nhanmaudk + "','"+Obj.Time_trabcpt_dk+"'," + Obj.TempMaMau + ",'','')";
           return ld.AddData(sql);
       }

       // add một mẫu cùng thông tin mẫu nhưng khác chỉ tiêu
       public bool addDataNew(Phieu_Nhan_Mau_Obj Obj)
       {

           string sql = "insert into PHIEU_YCPT(MAHD, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, TTKHCUNGCAP, NGUOINHANMAU, SEAL, TIME_NHANMAU_DK, MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, DONGIA, MAPP, TENPP,TRANGTHAI_CHUYEN,temp_MAMAU,TIME_TRABCPT_DK,MAQC,TENQC,NHOMCHITIEU,SOLUONGKIEM, MACD, TENCD,JOBNO,MAMAU) VALUES ( N'" + Obj.Mahd + "', N'" + Obj.Tenmau + "', N'" + Obj.Motamau + "', N'" + Obj.Nguonmau + "', N'" + Obj.Khoiluongmau + "', N'" + Obj.Tinhtrangmau + "', N'" + Obj.Ttkhcungcap + "',N'" + Obj.Nguoinhanmau + "', N'" + Obj.Seal + "', '" + Obj.Time_nhanmaudk + "', N'" + Obj.Mact + "', N'" + Obj.Tenct + "', " + Obj.Manenmau + ", N'" + Obj.Tennenmau + "', N'" + Obj.Lod + "', " + Obj.Madv + ", N'" + Obj.Tendv + "', " + Obj.Dongia + ", " + Obj.Mapp + ", N'" + Obj.Tenpp + "',N'Chưa chuyển'," + Obj.TempMaMau + ",'" + Obj.Time_trabcpt_dk + "'," + Obj.Maqc + ",N'" + Obj.Tenqc + "',N'"+Obj.Nhomct+"',1, N'"+Obj.Macd+"', N'"+Obj.Tencd+"','','')";
           return ld.AddData(sql);
       }

       // add một mẫu cùng thông tin mẫu cùng chỉ định
       public bool addData_ChiDinh(Phieu_Nhan_Mau_Obj Obj)
       {
           string sql = "insert into PHIEU_YCPT(MAHD, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, TTKHCUNGCAP, NGUOINHANMAU, SEAL, TIME_NHANMAU_DK, MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, MACD, TENCD,TRANGTHAI_CHUYEN) VALUES ( N'" + Obj.Mahd + "', N'" + Obj.Tenmau + "', N'" + Obj.Motamau + "', N'" + Obj.Nguonmau + "', N'" + Obj.Khoiluongmau + "', N'" + Obj.Tinhtrangmau + "', N'" + Obj.Ttkhcungcap + "',N'" + Obj.Nguoinhanmau + "', N'" + Obj.Seal + "', N'" + Obj.Time_nhanmaudk + "', N'" + Obj.Mact + "', N'" + Obj.Tenct + "', N'" + Obj.Manenmau + "', N'" + Obj.Tennenmau + "', N'" + Obj.Lod + "', N'" + Obj.Madv + "', N'" + Obj.Tendv + "', N'" + Obj.Macd + "', N'" + Obj.Tencd + "',N'Chưa chuyển')";
           return ld.AddData(sql);
       }

       public bool addChiDinh(Phieu_Nhan_Mau_Obj Obj)
       {
           string sql = "";
           return ld.AddData(sql);
       }

       // update chỉ tiêu đầu tiên
       public bool updateChiTieuDau(Phieu_Nhan_Mau_Obj Obj)
       {
           string sql = "update PHIEU_YCPT set SOLUONGKIEM=1, TRANGTHAI_CHUYEN=N'Chưa chuyển', MACT = N'" + Obj.Mact + "', TENCT = N'" + Obj.Tenct + "', MANENMAU = " + Obj.Manenmau + ", TENNENMAU = N'" + Obj.Tennenmau + "', LOD = N'" + Obj.Lod + "', MADV = " + Obj.Madv + ", TENDV = N'" + Obj.Tendv + "', DONGIA = " + Obj.Dongia + ", MAPP = " + Obj.Mapp + ", TENPP = '" + Obj.Tenpp + "',MAQC="+Obj.Maqc+",TENQC=N'"+Obj.Tenqc+"',NHOMCHITIEU=N'"+Obj.Nhomct+"', MACD = N'"+Obj.Macd+"', TENCD= N'"+Obj.Tencd+"' where MAPHIEU_YCPT = " + Obj.Maphieu;
           return ld.AddData(sql);
       }

       // update chỉ tiêu của chỉ định đầu tiên
       public bool updateChiDinhDau(Phieu_Nhan_Mau_Obj Obj)
       {
           string sql = "update PHIEU_YCPT set TRANGTHAI_CHUYEN=N'Chưa chuyển', MACT = N'" + Obj.Mact + "', TENCT = N'" + Obj.Tenct + "', MANENMAU = N'" + Obj.Manenmau + "', TENNENMAU = N'" + Obj.Tennenmau + "', LOD = N'" + Obj.Lod + "', MADV = N'" + Obj.Madv + "', TENDV = N'" + Obj.Tendv + "', MACD = N'" + Obj.Macd + "', TENCD = N'" + Obj.Tencd + "' where MAPHIEU_YCPT = N'" + Obj.Maphieu + "' ";
           return ld.AddData(sql);
       }

       // lấy lại thông tin mẫu
       public DataTable CopyMau(string temp_MAMAU)
       {
           string sql = "select MAPHIEU_YCPT,TIME_NHANMAU_DK,MAHD,TENMAU,MOTAMAU,KHOILUONGMAU,TINHTRANGMAU,TTKHCUNGCAP,NGUONMAU,SEAL,NGUOINHANMAU,temp_MAMAU,MACT,TIME_TRABCPT_DK from PHIEU_YCPT where temp_MAMAU = " + temp_MAMAU;
           return ld.GetData(sql);
       }

       // kiểm tra nên update hay add
       public DataTable Check_UOA(Phieu_Nhan_Mau_Obj Obj)
       {
           string sql = "select MACT from PHIEU_YCPT where MAPHIEU_YCPT =  '"+Obj.Maphieu+"'";
           return ld.GetData(sql);
       }



       ////////////Lấy mã mẫu tạm
       public string getTempMaMau()
       {
           return ld.GetData("exec sp_temp_MaMau").Rows[0][0].ToString();
       }

    }
}
