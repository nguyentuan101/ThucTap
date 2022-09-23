using QuanLyMauKiemNghiem.Object;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
    public class CapNhatMau_Mod
    {
        LoadData ld = new LoadData();

        // lấy danh sách mẫu 
        public DataTable GetDSMau(bool chuacoCT)
        {

            string sql = "select distinct c.MAKH+'-'+c.TENKH as TENKH,b.MAHD, b.MAHD +' '+TENHD as tenhd, JOBNO, MAMAU, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, SEAL, TIME_NHANMAU_DK, CONVERT(date,TIME_NHANMAU_TT,101) as TIME_NHANMAU_TT, TTKHCUNGCAP, NGUOINHANMAU, TIME_TRABCPT_DK, CONVERT(date,TIME_TRABCPT_TT,101) as TIME_TRABCPT_TT,TIME_CHUYENMAU,temp_MAMAU,case TRANGTHAI_CHUYEN when N'Kiểm lại' then N'Kiểm lại' when N'Chưa chuyển' then N'Chưa chuyển TTKN' when N'Đã chuyển' then N'TTKN đã nhận' when N'Hoàn thành' then N'Đã có KQPT' when N'Đã xuất' then N'Đã xuất KQPT' else N'Chưa có chỉ tiêu' end as TRANGTHAI_CHUYEN from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on b.MAKH=c.MAKH";
            if (chuacoCT == true)
                sql = "select distinct c.MAKH+'-'+c.TENKH as TENKH,b.MAHD, b.MAHD +' '+TENHD as tenhd, JOBNO, MAMAU, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, SEAL, TIME_NHANMAU_DK, CONVERT(date,TIME_NHANMAU_TT,101) as TIME_NHANMAU_TT, TTKHCUNGCAP, NGUOINHANMAU, TIME_TRABCPT_DK, CONVERT(date,TIME_TRABCPT_TT,101) as TIME_TRABCPT_TT,TIME_CHUYENMAU,temp_MAMAU,case TRANGTHAI_CHUYEN when N'Kiểm lại' then N'Kiểm lại' when N'Chưa chuyển' then N'Chưa chuyển TTKN' when N'Đã chuyển' then N'TTKN đã nhận' when N'Hoàn thành' then N'Đã có KQPT' when N'Đã xuất' then N'Đã xuất KQPT' else N'Chưa có chỉ tiêu' end as TRANGTHAI_CHUYEN from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD join KHACH_HANG c on b.MAKH=c.MAKH where MACT is null or MACT=''";
            return ld.GetData(sql);
        }
        
        // lấy danh sách mẫu 
        public DataTable GetNoiKiem()
        {
            string sql = "select * from NOI_KIEM";
            return ld.GetData(sql);
        }
        public DataTable GetPhuongPhap()
        {
            string sql = "select * from PHUONG_PHAP";
            return ld.GetData(sql);
        }


        // lấy danh chi tiêu của mẫu
        public DataTable GetDSChiTieu(string tempMaMau)
        {
            //string sql = "select MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, MAPP, TENPP, DONGIA, MANK, TENNK, MAPHIEU_YCPT from PHIEU_YCPT where temp_MAMAU='"+tempMaMau+"'";
            string sql = "select *,case TRANGTHAI_CHUYEN when N'Kiểm lại' then N'Kiểm lại' when N'Chưa chuyển' then N'Chưa chuyển TTKN' when N'Đã chuyển' then N'TTKN đã nhận' when N'Hoàn thành' then N'Đã có KQPT' when N'Đã xuất' then N'Đã xuất KQPT' else N'Chưa có chỉ tiêu' end as TRANGTHAI_CHUYENCT from PHIEU_YCPT where temp_MAMAU='" + tempMaMau + "'";
            return ld.GetData(sql);
        }

        // lấy danh sách chỉ tiêu
        public DataTable GetChiTieu(string tempMaMau)
        {
            string sql = "select ct.MACT, ct.TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, gct.GIA, MAPP, TENPP,NHOMCHITIEU,MAQC,TENQC from CHI_TIEU as ct left join GIA_CHI_TIEU as gct on ct.MACT = gct.MACT where ct.MACT NOT IN (select ISNULL(MACT,'') as MACT from PHIEU_YCPT where temp_MAMAU='" + tempMaMau + "')";
            return ld.GetData(sql);
        }

        // cập nhật thông tin mẫu
        public bool Update(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "update PHIEU_YCPT set MAHD=N'" + Obj.Mahd + "', JOBNO=N'" + Obj.Jobno + "', MAMAU=N'" + Obj.Mamau + "', TENMAU=N'" + Obj.Tenmau + "', MOTAMAU=N'" + Obj.Motamau + "', KHOILUONGMAU=N'" + Obj.Khoiluongmau + "', TINHTRANGMAU=N'" + Obj.Tinhtrangmau + "', SEAL=N'" + Obj.Seal + "', TIME_NHANMAU_DK=N'" + Obj.Time_nhanmaudk + "', TTKHCUNGCAP=N'" + Obj.Ttkhcungcap + "', NGUOINHANMAU=N'" + Obj.Nguoinhanmau + "', TIME_TRABCPT_DK='" + Obj.Time_trabcpt_dk + "' where temp_MAMAU='"+Obj.TempMaMau+"'";
            return ld.AddData(sql);
        }
        // cập nhật nơi kiểm cho chỉ tiêu
        public bool delete(string  temp)
        {
            string sql = "delete PHIEU_YCPT where temp_MAMAU='"+temp+"'";
            return ld.AddData(sql);
        }
        // lấy thông tin của mẫu cần cập nhật chỉ tiêu
        public DataTable GetData(string tempMAMAU)
        {
            string sql = "select phieu.MAHD, kh.TENKH, JOBNO, MAMAU, TENMAU, DONGIA from PHIEU_YCPT as phieu, HOP_DONG as hd, KHACH_HANG as kh where phieu.MAHD = hd.MAHD and kh.MAKH = hd.MAKH and temp_MAMAU=" + tempMAMAU + "";
            return ld.GetData(sql);
        }

        // lấy danh sách chỉ định
        public DataTable GetChiDinh()
        {
            string sql = "select distinct MACD, TENCD from CHI_DINH_CHI_TIEU as ctcd, CHI_TIEU as ct where ctcd.MACT = ct.MACT";
            return ld.GetData(sql);
        }

        // lấy danh sách chỉ tiêu theo chỉ định
        public DataTable GetChiTieu_ChiDinh(string macd, string manenmau)
        {
            string sql = "select ctcd.MACT, ctcd.TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, MACD, TENCD, GIA, MAPP, TENPP from CHI_DINH_CHI_TIEU as ctcd, CHI_TIEU as ct, GIA_CHI_TIEU as gct where ctcd.MACT = ct.MACT and gct.MACT = ct.MACT and MACD = '" + macd + "' and MANENMAU = '" + manenmau + "'";
            return ld.GetData(sql);
        }

        // lấy danh sách loại giá
        public DataTable GetLoaiGia(string mact)
        {
            string sql = "select TENPHANLOAIGIA,GIA from CHI_TIEU a join GIA_CHI_TIEU b on a.MACT=b.MACT where b.MACT='"+mact+"'";
            return ld.GetData(sql);
        }

        // cập nhật nơi kiểm cho chỉ tiêu
        public bool UpdateNoiKiem(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "update PHIEU_YCPT set MANK = N'" + Obj.Mank + "', TENNK = N'" + Obj.Tennk + "' where MAPHIEU_YCPT='" + Obj.Maphieu + "'";
            return ld.AddData(sql);
        }

        // cập nhật đơn giá cho chỉ tiêu
        public bool UpdateDongia(Gia_Chi_Tieu_Obj Obj)
        {
            string sql = "update PHIEU_YCPT set DONGIA = N'"+Obj.Gia+"' where MAPHIEU_YCPT='"+Obj.Magia+"'";
            return ld.AddData(sql);
        }


        // update chỉ tiêu đầu tiên
        public bool updateChiTieuDau(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "update PHIEU_YCPT set SOLUONGKIEM=1, TRANGTHAI_CHUYEN=N'Chưa chuyển', MACT = N'" + Obj.Mact + "', TENCT = N'" + Obj.Tenct + "', MANENMAU = " + Obj.Manenmau + ", TENNENMAU = N'" + Obj.Tennenmau + "', LOD = N'" + Obj.Lod + "', MADV = " + Obj.Madv + ", TENDV = N'" + Obj.Tendv + "', DONGIA = " + Obj.Dongia + ", MAPP = " + Obj.Mapp + ", TENPP = '" + Obj.Tenpp + "',MAQC=" + Obj.Maqc + ",TENQC=N'" + Obj.Tenqc + "',NHOMCHITIEU=N'" + Obj.Nhomct + "' where MAPHIEU_YCPT = " + Obj.Maphieu;
            return ld.AddData(sql);
        }

        // update chỉ tiêu của chỉ định đầu tiên
        public bool updateChiDinhDau(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "update PHIEU_YCPT set TRANGTHAI_CHUYEN=N'Chưa chuyển', MACT = N'" + Obj.Mact + "', TENCT = N'" + Obj.Tenct + "', MANENMAU = N'" + Obj.Manenmau + "', TENNENMAU = N'" + Obj.Tennenmau + "', LOD = N'" + Obj.Lod + "', MADV = N'" + Obj.Madv + "', TENDV = N'" + Obj.Tendv + "', MACD = N'" + Obj.Macd + "', TENCD = N'" + Obj.Tencd + "', DONGIA = '" + Obj.Dongia + "', MAPP = '" + Obj.Mapp + "', TENPP = '" + Obj.Tenpp + "' where MAHD='" + Obj.Mahd + "' and TENMAU=N'" + Obj.Tenmau + "' and TIME_NHANMAU_DK='" + Obj.Time_nhanmaudk + "' ";
            return ld.AddData(sql);
        }

        // lấy lại thông tin mẫu
        public DataTable CopyMau(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "select MAHD, TENMAU, SEAL, KHOILUONGMAU, TINHTRANGMAU, MOTAMAU, NGUOINHANMAU, TTKHCUNGCAP, NGUONMAU, TIME_NHANMAU_DK from PHIEU_YCPT where MAHD='" + Obj.Mahd + "' and TENMAU=N'" + Obj.Tenmau + "' and TIME_NHANMAU_DK='" + Obj.Time_nhanmaudk + "'";
            return ld.GetData(sql);
        }




        // add một mẫu cùng thông tin mẫu cùng chỉ định
        public bool addData_ChiDinh(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "insert into PHIEU_YCPT(MAHD, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, TTKHCUNGCAP, NGUOINHANMAU, SEAL, TIME_NHANMAU_DK, MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, MACD, TENCD, DONGIA, MAPP, TENPP,TRANGTHAI_CHUYEN) VALUES ( N'" + Obj.Mahd + "', N'" + Obj.Tenmau + "', N'" + Obj.Motamau + "', N'" + Obj.Nguonmau + "', N'" + Obj.Khoiluongmau + "', N'" + Obj.Tinhtrangmau + "', N'" + Obj.Ttkhcungcap + "',N'" + Obj.Nguoinhanmau + "', N'" + Obj.Seal + "', N'" + Obj.Time_nhanmaudk + "', N'" + Obj.Mact + "', N'" + Obj.Tenct + "', N'" + Obj.Manenmau + "', N'" + Obj.Tennenmau + "', N'" + Obj.Lod + "', N'" + Obj.Madv + "', N'" + Obj.Tendv + "', N'" + Obj.Macd + "', N'" + Obj.Tencd + "', N'" + Obj.Dongia + "', N'" + Obj.Mapp + "', N'" + Obj.Tenpp + "',N'Chưa chuyển')";
            return ld.AddData(sql);
        }


        // add một mẫu cùng thông tin mẫu nhưng khách chỉ tiêu
        public bool addDataNew(Phieu_Nhan_Mau_Obj Obj)
        {
            string sql = "insert into PHIEU_YCPT(MAHD, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, TTKHCUNGCAP, NGUOINHANMAU, SEAL, TIME_NHANMAU_DK, MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, DONGIA, MAPP, TENPP,TRANGTHAI_CHUYEN,temp_MAMAU,TIME_NHANMAU_TT,TIME_TRABCPT_DK,JOBNO,MAMAU,MAQC,TENQC,SOLUONGKIEM) VALUES ( N'" + Obj.Mahd + "', N'" + Obj.Tenmau + "', N'" + Obj.Motamau + "', N'" + Obj.Nguonmau + "', N'" + Obj.Khoiluongmau + "', N'" + Obj.Tinhtrangmau + "', N'" + Obj.Ttkhcungcap + "',N'" + Obj.Nguoinhanmau + "', N'" + Obj.Seal + "', '" + Obj.Time_nhanmaudk + "', N'" + Obj.Mact + "', N'" + Obj.Tenct + "', " + Obj.Manenmau + ", N'" + Obj.Tennenmau + "', N'" + Obj.Lod + "', " + Obj.Madv + ", N'" + Obj.Tendv + "'," + Obj.Dongia + ", " + Obj.Mapp + ", N'" + Obj.Tenpp + "',N'Chưa chuyển'," + Obj.TempMaMau + ",'" + Obj.Time_nhanmau_tt + "','" + Obj.Time_trabcpt_dk + "','" + Obj.Jobno + "','" + Obj.Mamau + "'," + Obj.Maqc + ",'" + Obj.Tenqc + "',1)";
            if(Obj.Time_nhanmau_tt=="")
                sql = "insert into PHIEU_YCPT(MAHD, TENMAU, MOTAMAU, NGUONMAU, KHOILUONGMAU, TINHTRANGMAU, TTKHCUNGCAP, NGUOINHANMAU, SEAL, TIME_NHANMAU_DK, MACT, TENCT, MANENMAU, TENNENMAU, LOD, MADV, TENDV, DONGIA, MAPP, TENPP,TRANGTHAI_CHUYEN,temp_MAMAU,TIME_TRABCPT_DK,JOBNO,MAMAU,MAQC,TENQC,SOLUONGKIEM) VALUES ( N'" + Obj.Mahd + "', N'" + Obj.Tenmau + "', N'" + Obj.Motamau + "', N'" + Obj.Nguonmau + "', N'" + Obj.Khoiluongmau + "', N'" + Obj.Tinhtrangmau + "', N'" + Obj.Ttkhcungcap + "',N'" + Obj.Nguoinhanmau + "', N'" + Obj.Seal + "', '" + Obj.Time_nhanmaudk + "', N'" + Obj.Mact + "', N'" + Obj.Tenct + "', " + Obj.Manenmau + ", N'" + Obj.Tennenmau + "', N'" + Obj.Lod + "', " + Obj.Madv + ", N'" + Obj.Tendv + "'," + Obj.Dongia + ", " + Obj.Mapp + ", N'" + Obj.Tenpp + "',N'Chưa chuyển'," + Obj.TempMaMau + ",'" + Obj.Time_trabcpt_dk + "','" + Obj.Jobno + "','" + Obj.Mamau + "'," + Obj.Maqc + ",'" + Obj.Tenqc + "',1)";
            return ld.AddData(sql);
        }

        // add một giá của chỉ tiêu vào bảng giá chỉ tiêu
        public bool AddGiaChiTieu(Gia_Chi_Tieu_Obj Obj)
        {
            string sql = "insert into GIA_CHI_TIEU(MACT, TENCT, GIA, TENPHANLOAIGIA) values (N'"+Obj.Mact+"', N'"+Obj.Tenct+"', N'"+Obj.Gia+"', N'"+Obj.Tenphanloaigia+"')";
            return ld.AddData(sql);
        }

        // check giá chỉ tiêu
        public DataTable Check_Gia(Gia_Chi_Tieu_Obj Obj)
        {
            string sql = "select * from gia_chi_tieu where MACT = '" + Obj.Mact + "' and TENPHANLOAIGIA='"+Obj.Tenphanloaigia+"'";
            return ld.GetData(sql);
        }

        public bool CapNhatChiTieuKiem(Gia_Chi_Tieu_Obj Obj, string mamau)
        {
            string sql = "update PHIEU_YCPT set MAPP=" + Obj.Mapp + ",TENPP=N'" + Obj.Tenpp + "',MANK=" + Obj.Mank + ",TENNK=N'" + Obj.Tennk + "',DONGIA=" + Obj.Gia + " where MAPHIEU_YCPT='" + Obj.Maphieu + "';update CONG_NO set MAPP=" + Obj.Mapp + ",TENPP=N'" + Obj.Tenpp + "',MANK=" + Obj.Mank + ",TENNK=N'" + Obj.Tennk + "',DONGIA=" + Obj.Gia + " where MAMAU = '"+mamau+"' and MACT = '"+Obj.Mact+"';";
            return ld.AddData(sql);
        }

        
        public bool XoaChiTieuKiem(string maphieu)
        {
            string sql = "delete PHIEU_YCPT where MAPHIEU_YCPT=" + maphieu;
            return ld.AddData(sql);
        }

        public bool YeuCauKiemLai(string maphieu)
        {
            string sql = "update PHIEU_YCPT set TRANGTHAI_CHUYEN=N'Kiểm lại', TRANGTHAI_XUAT=N'Kiểm lại',TRANGTHAI_KQ=N'Kiểm lại' where MAPHIEU_YCPT=" + maphieu;
            return ld.AddData(sql);
        }
    }
}
