using QuanLyMauKiemNghiem.Object;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
    public class CongNoChuaKQ_Mod
    {
        LoadData ld = new LoadData();

        // lấy danh sách khách hàng cho combobox khách hàng
        public DataTable GetKH()
        {
            string sql = "SELECT DISTINCT C.MAKH,C.TENKH FROM PHIEU_YCPT  A JOIN HOP_DONG B ON A.MAHD=B.MAHD JOIN KHACH_HANG C ON B.MAKH=C.MAKH and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        // lấy danh sách khách hàng theo hợp đồng cho combobox khách hàng
        public DataTable GetKH(string mahd)
        {
            string sql = "SELECT DISTINCT C.MAKH,C.TENKH FROM PHIEU_YCPT  A JOIN HOP_DONG B ON A.MAHD=B.MAHD JOIN KHACH_HANG C ON B.MAKH=C.MAKH WHERE A.MAHD= '" + mahd + "' and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        // lấy danh sách khách hàng theo jobno cho combobox khách hàng
        public DataTable GetKH1(string jobno)
        {
            string sql = "SELECT DISTINCT C.MAKH,C.TENKH FROM PHIEU_YCPT  A JOIN HOP_DONG B ON A.MAHD=B.MAHD JOIN KHACH_HANG C ON B.MAKH=C.MAKH WHERE JOBNO= '" + jobno + "' and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        // lấy danh sách hợp đồng cho combobox hợp đồng
        public DataTable GetHD()
        {
            string sql = "select distinct phieu.MAHD, TENHD from PHIEU_YCPT as phieu left join HOP_DONG on HOP_DONG.MAHD = phieu.MAHD where (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        // lấy danh sách hợp đồng theo khách hàng cho combobox hợp đồng
        public DataTable GetHD(string makh)
        {
            string sql = "select distinct phieu.MAHD, TENHD from PHIEU_YCPT as phieu left join HOP_DONG on HOP_DONG.MAHD = phieu.MAHD where HOP_DONG.MAKH='" + makh + "' and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        // lấy danh sách hợp đồng theo jobno cho combobox hợp đồng
        public DataTable GetHD1(string jobno)
        {
            string sql = "select distinct phieu.MAHD, TENHD from PHIEU_YCPT as phieu left join HOP_DONG on HOP_DONG.MAHD = phieu.MAHD where phieu.JOBNO='" + jobno + "' and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        // lấy danh sách jobno cho combobox jobno
        public DataTable GetJobNo()
        {
            string sql = "select distinct JOBNO from PHIEU_YCPT where (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }
        // lấy danh sách jobno theo khách hàng cho combobox jobno
        public DataTable GetJobNo(string makh)
        {
            string sql = "select distinct JOBNO from PHIEU_YCPT a join HOP_DONG b on a.MAHD=b.MAHD where MAKH = N'" + makh + "' and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }
        // lấy danh sách jobno theo hợp đồng cho combobox jobno
        public DataTable GetJobNo1(string mahd)
        {
            string sql = "select distinct JOBNO from PHIEU_YCPT where MAHD = N'" + mahd + "' and (JOBNO is not null and JOBNO !=  N'')";
            return ld.GetData(sql);
        }

        public DataTable GetData(Phieu_Nhan_Mau_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if (cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONGKIEM, DONGIA, (SOLUONGKIEM * DONGIA) as TONGTIEN, TENCD, TENKH, MAPHIEU_YCPT from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONGKIEM, DONGIA, (SOLUONGKIEM * DONGIA) as TONGTIEN, TENCD, TENKH, MAPHIEU_YCPT from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD  where hd.makh = N'" + cnObj.Makh + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else
                {
                    string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONGKIEM, DONGIA, (SOLUONGKIEM * DONGIA) as TONGTIEN, TENCD, TENKH, MAPHIEU_YCPT from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where hd.MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONGKIEM, DONGIA, (SOLUONGKIEM * DONGIA) as TONGTIEN, TENCD, TENKH, MAPHIEU_YCPT from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD  where hd.MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "'";
                return ld.GetData(sql);
            }
        }


        // lấy danh sách chỉ tiêu
        public DataTable GetCT(Phieu_Nhan_Mau_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if (cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        string sql = "select DISTINCT MACT,TENCT, DONGIA from PHIEU_YCPT as cn where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select DISTINCT MACT,TENCT, DONGIA from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD  where hd.makh = N'" + cnObj.Makh + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else
                {
                    string sql = "select DISTINCT MACT,TENCT, DONGIA from PHIEU_YCPT as cn where cn.mahd = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select DISTINCT MACT,TENCT, DONGIA from PHIEU_YCPT as cn where MAHD = N'"+cnObj.Mahd+"' and JOBNO = '"+cnObj.Jobno+"'";
                return ld.GetData(sql);
            }
        }


        // lấy danh sách tên mẫu
        public DataTable GetTM(Phieu_Nhan_Mau_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if (cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        string sql = "select distinct TENMAU , MAMAU  from PHIEU_YCPT as cn where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select distinct TENMAU , MAMAU  from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD  where hd.makh = N'" + cnObj.Makh + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else
                {
                    string sql = "select distinct TENMAU , MAMAU  from PHIEU_YCPT as cn where MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select distinct TENMAU , MAMAU  from PHIEU_YCPT as cn where MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "'";
                return ld.GetData(sql);
            }
        }


        public DataTable GetDCKH(string tenkh)
        {
            string sql = "select DIACHI from KHACH_HANG where TENKH = N'" + tenkh + "'";
            return ld.GetData(sql);
        }

        // lấy tổng số tiền của từng mẫu
        public DataTable GetTT(Phieu_Nhan_Mau_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            string sql = "select sum((SOLUONGKIEM * DONGIA)) as TONGTIEN from PHIEU_YCPT where  MAMAU = N'" + cnObj.Mamau + "'";
            return ld.GetData(sql);
        }

        public DataTable GetSL(Phieu_Nhan_Mau_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            string sql = "select SOLUONGKIEM from PHIEU_YCPT  as cn where cn.MAMAU = '" + cnObj.Mamau + "' and cn.MACT = N'" + cnObj.Mact + "' and cn.DONGIA = N'" + cnObj.Dongia + "'";
            return ld.GetData(sql);
        }

          // lấy tổng tiền của nguyên phiếu
        public DataTable GetTongTien(Phieu_Nhan_Mau_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if (cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        string sql = "select sum((SOLUONGKIEM * DONGIA)) as TONGTIEN from PHIEU_YCPT where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select sum((SOLUONGKIEM * DONGIA)) as TONGTIEN from PHIEU_YCPT as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD  where hd.makh = N'" + cnObj.Makh + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else
                {
                    string sql = "select sum((SOLUONGKIEM * DONGIA)) as TONGTIEN from PHIEU_YCPT where MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select sum((SOLUONGKIEM * DONGIA)) as TONGTIEN from PHIEU_YCPT where MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "'";
                return ld.GetData(sql);
            }
        }

    }
}
