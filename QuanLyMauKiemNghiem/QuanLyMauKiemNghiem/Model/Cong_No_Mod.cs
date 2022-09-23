using QuanLyMauKiemNghiem.Object;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Model
{
    public class Cong_No_Mod
    {
        LoadData ld = new LoadData();

        // lấy danh sách khách hàng cho combobox khách hàng
        public DataTable GetKH()
        {
            string sql = "select distinct CONG_NO.MAKH, KHACH_HANG.TENKH from CONG_NO left join KHACH_HANG on KHACH_HANG.MAKH = CONG_NO.MAKH";
            return ld.GetData(sql);
        }

        // lấy danh sách khách hàng theo hợp đồng cho combobox khách hàng
        public DataTable GetKH(string mahd)
        {
            string sql = "SELECT DISTINCT C.MAKH,C.TENKH FROM CONG_NO  A JOIN HOP_DONG B ON A.MAHD=B.MAHD JOIN KHACH_HANG C ON B.MAKH=C.MAKH where A.MAHD = '" + mahd + "'";
            return ld.GetData(sql);
        }

        // lấy danh sách khách hàng theo jobno cho combobox khách hàng
        public DataTable GetKH1(string jobno)
        {
            string sql = "SELECT DISTINCT C.MAKH,C.TENKH FROM CONG_NO  A JOIN HOP_DONG B ON A.MAHD=B.MAHD JOIN KHACH_HANG C ON B.MAKH=C.MAKH where A.JOBNO = '" + jobno + "'";
            return ld.GetData(sql);
        }

        // lấy danh sách hợp đồng cho combobox hợp đồng
        public DataTable GetHD()
        {
            string sql = "select distinct CONG_NO.MAHD, TENHD from CONG_NO left join HOP_DONG on HOP_DONG.MAHD = CONG_NO.MAHD";
            return ld.GetData(sql);
        }

        // lấy danh sách hợp đồng theo khách hàng cho combobox hợp đồng
        public DataTable GetHD(string makh)
        {
            string sql = "select distinct CONG_NO.MAHD, TENHD from CONG_NO left join HOP_DONG on HOP_DONG.MAHD = CONG_NO.MAHD where CONG_NO.MAKH = N'" + makh + "'";
            return ld.GetData(sql);
        }

        // lấy danh sách hợp đồng theo jobno cho combobox hợp đồng
        public DataTable GetHD1(string jobno)
        {
            string sql = "select distinct CONG_NO.MAHD, TENHD from CONG_NO left join HOP_DONG on HOP_DONG.MAHD = CONG_NO.MAHD where CONG_NO.JOBNO = N'" + jobno + "'";
            return ld.GetData(sql);
        }

        public DataTable GetJobNo()
        {
            string sql = "select distinct JOBNO from CONG_NO";
            return ld.GetData(sql);
        }

        public DataTable GetJobNo(string makh)
        {
            string sql = "select distinct JOBNO from CONG_NO where MAKH = N'"+makh+"'";
            return ld.GetData(sql);
        }

        public DataTable GetJobNo1(string mahd)
        {
            string sql = "select distinct JOBNO from CONG_NO where MAHD = N'" + mahd + "'";
            return ld.GetData(sql);
        }

        public DataTable GetData(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if(cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        if(thoigian_bd == "0001/01/01" && thoigian_kt == "0001/01/01")
                        {
                            string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD";
                            return ld.GetData(sql);
                        }
                        else
                        {
                            string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                            return ld.GetData(sql);
                        }

                    }
                    else if (thoigian_bd == "0001/01/01" && thoigian_kt == "0001/01/01")
                    {
                        string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where cn.makh = N'" + cnObj.Makh + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where cn.makh = N'" + cnObj.Makh + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else if (thoigian_bd == "0001/01/01" && thoigian_kt == "0001/01/01")
                {
                    string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where cn.MAHD = N'" + cnObj.Mahd + "'";
                    return ld.GetData(sql);
                }

                else
                {
                    string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where cn.MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }

            else
            {
                string sql = "select JOBNO, MAMAU, TENMAU, MACT, TENCT, SOLUONG, DONGIA, (SOLUONG * DONGIA) as TONGTIEN, TENCD, hd.TENKH, MACNO, TENNK, cn.MAHD from CONG_NO as cn left join HOP_DONG as hd on hd.MAHD = cn.MAHD where cn.MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "'";
                return ld.GetData(sql);
            }

        }

        public DataTable GetDCKH(string tenkh)
        {
            string sql = "select DIACHI from KHACH_HANG where TENKH = N'"+tenkh+"'";
            return ld.GetData(sql);
        }


        // lấy tổng số tiền của từng mẫu
        public DataTable GetTT(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO where  MAMAU = N'"+cnObj.Mamau+"'";
            return ld.GetData(sql);
        }

        public DataTable GetSL(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            string sql = "select SOLUONG from CONG_NO  as cn where cn.MAMAU = '" + cnObj.Mamau + "' and cn.MACT = N'" + cnObj.Machitieu + "' and cn.DONGIA = N'" + cnObj.Dongia + "'";
            return ld.GetData(sql);
        }

        // lấy tổng tiền của nguyên phiếu
        public DataTable GetTongTien(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if(cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        if (thoigian_bd == "0001/01/01" && thoigian_kt == "0001/01/01")
                        {
                            string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO";
                            return ld.GetData(sql);
                        }
                        else
                        {
                            string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                            return ld.GetData(sql);
                        }
                    }
                    else if (thoigian_bd == "0001/01/01" && thoigian_kt == "0001/01/01")
                    {
                        string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO  where makh = N'" + cnObj.Makh + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO  where makh = N'"+cnObj.Makh+"' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else if (thoigian_bd == "0001/01/01" && thoigian_kt == "0001/01/01")
                {
                    string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO where MAHD = N'" + cnObj.Mahd + "'";
                    return ld.GetData(sql);
                }

                else
                {
                    string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO where MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select sum((SOLUONG * DONGIA)) as TONGTIEN from CONG_NO where MAHD = N'" + cnObj.Mahd + "'";
                return ld.GetData(sql);
            }
        }



        // lấy danh sách chỉ tiêu
        public DataTable GetCT(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if (cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        string sql = "select DISTINCT MACT,TENCT, DONGIA from CONG_NO as cn where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select DISTINCT MACT,TENCT, DONGIA from CONG_NO as cn where cn.makh = N'"+cnObj.Makh+"' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else
                {
                    string sql = "select DISTINCT MACT,TENCT, DONGIA from CONG_NO as cn where MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select DISTINCT MACT,TENCT, DONGIA from CONG_NO as cn where MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "'";
                return ld.GetData(sql);
            }
        }


        // lấy danh sách tên mẫu
        public DataTable GetTM(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        {
            if (cnObj.Jobno == null)
            {
                if (cnObj.Mahd == null)
                {
                    if (cnObj.Makh == null)
                    {
                        string sql = "select distinct TENMAU , MAMAU  from CONG_NO as cn where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                    else
                    {
                        string sql = "select distinct TENMAU , MAMAU  from CONG_NO as cn where cn.makh = N'" + cnObj.Makh + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                        return ld.GetData(sql);
                    }
                }
                else
                {
                    string sql = "select distinct TENMAU , MAMAU  from CONG_NO as cn where MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
                    return ld.GetData(sql);
                }

            }
            else
            {
                string sql = "select distinct TENMAU , MAMAU  from CONG_NO as cn where MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "'";
                return ld.GetData(sql);
            }
        }

        // cập nhật đơn giá cho chỉ tiêu
        public bool updateDonGia(Cong_No_Obj Obj)
        {
            string sql = "UPDATE CONG_NO SET DONGIA = '" + Obj.Dongia + "' WHERE MACNO = '" + Obj.Macno + "'";
            return ld.AddData(sql);
        }


        //// cập nhật ngày thanh toán
        //public bool updateNgayThanhToan(Cong_No_Obj cnObj, string thoigian_bd, string thoigian_kt)
        //{
        //    if (cnObj.Jobno == null)
        //    {
        //        if (cnObj.Mahd == null)
        //        {
        //            if (cnObj.Makh == null)
        //            {
        //                string sql = "UPDATE CONG_NO SET NGAYTHANHTOAN = '" + cnObj.Ngaythanhtoan + "' where TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
        //                return ld.AddData(sql);
        //            }
        //            else
        //            {
        //                string sql = "UPDATE CONG_NO SET NGAYTHANHTOAN = '" + cnObj.Ngaythanhtoan + "' where MAKH = '"+cnObj.Makh+"' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
        //                return ld.AddData(sql);
        //            }
        //        }
        //        else
        //        {
        //            string sql = "UPDATE CONG_NO SET NGAYTHANHTOAN = '" + cnObj.Ngaythanhtoan + "' where MAHD = N'" + cnObj.Mahd + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
        //            return ld.AddData(sql);
        //        }

        //    }
        //    else
        //    {
        //        if (cnObj.Mahd == null)
        //        {
        //            string sql = "UPDATE CONG_NO SET NGAYTHANHTOAN = '" + cnObj.Ngaythanhtoan + "' where JOBNO = '" + cnObj.Jobno + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
        //            return ld.AddData(sql);
        //        }
        //        else
        //        {
        //            string sql = "UPDATE CONG_NO SET NGAYTHANHTOAN = '" + cnObj.Ngaythanhtoan + "' where MAHD = N'" + cnObj.Mahd + "' and JOBNO = '" + cnObj.Jobno + "' and TIME_NHANMAU_TT between '" + thoigian_bd + "' and '" + thoigian_kt + "'";
        //            return ld.AddData(sql);
        //        }

        //    }
        //}

        //iNSERT CONG NO
        public bool insertCongNo(Cong_No_Obj Obj)
        {
            string sql = "insert into CONG_NO values (" + Obj.Maphieu + ",'" + Obj.Mahd + "','" + Obj.Jobno + "','" + Obj.Mamau + "',N'" + Obj.Tenmau + "','" + Obj.Time_nhanmau_tt + "'," + Obj.Manenmau + ",N'" + Obj.Nenmau + "'," + Obj.Machidinh + ",N'" + Obj.Chidinh + "','" + Obj.Machitieu + "',N'" + Obj.Chitieu + "'," + Obj.Maqc + ",N'" + Obj.Tenqc + "'," + Obj.Maphuongphap + ",N'" + Obj.Phuongphap + "'," + Obj.Dongia + "," + Obj.Lankiemthu + ",'" + Obj.Soluong + "',GETDATE(),'" + Obj.Lod + "'," + Obj.Madv + ",N'" + Obj.Tendv + "','" + Obj.Nhomchitieu + "','" + Obj.Userxuat + "','" + Obj.Makh + "',N'" + Obj.Tenkh + "',N'" + Obj.Ketqua + "',"+Obj.Mank+",N'"+Obj.Tennk+"')";
            return ld.AddData(sql);
        }
        public string LanKiemThu(string maphieu)
        {
            return ld.GetData(" select count(LANKIEMTHU)+1 from CONG_NO where MAPHIEU_YCPT="+maphieu).Rows[0][0].ToString();
        }
    }
}
