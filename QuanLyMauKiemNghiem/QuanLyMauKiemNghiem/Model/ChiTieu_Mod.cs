using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using QuanLyMauKiemNghiem.Object;
namespace QuanLyMauKiemNghiem.Model
{
    class ChiTieu_Mod
    {
        LoadData ld = new LoadData();

        public string autoMaChiTieu()
        {
            string sql = "SELECT MAX(RIGHT(MACT,7))+1 FROM CHI_TIEU";
            DataTable dt = ld.GetData(sql);
            string ma = "";
            DataRow ROW = dt.Rows[0];

            if (ROW[0].ToString() == "")
                ma = ma + "0000000";
            else
            {
                string h;
                ma = "";
                // k = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][0].ToString().Substring(0, 7));
                h = ROW[0].ToString();
                int k = int.Parse(h);
                if (k < 10)
                    ma = ma + "000000";
                else
                {
                    if (k < 100)
                        ma = ma + "00000";
                    else
                    {
                        if (k < 1000)
                            ma = ma + "0000";
                        else
                        {
                            if (k < 10000)
                                ma = ma + "000";
                            else
                            {
                                if (k < 100000)
                                    ma = ma + "00";
                                else
                                    ma = ma + "0";
                            }
                        }
                    }
                }
                ma = ma + k.ToString();
            }
            return "CT-" + ma;
        }
        //Lấy nhóm chỉ tiêu
        public DataTable GetNhomChiTieu()
        {
            string sql = "select distinct NHOMCHITIEU from CHI_TIEU";
            return ld.GetData(sql);
        }
        //Kiểm tra gia chi tiêu
        public DataTable kiemtragiachitieu(string mact,string tenplg)
        {
            string sql = "select * from GIA_CHI_TIEU where MACT='"+mact+"' and TENPHANLOAIGIA=N'"+tenplg+"'";
            return ld.GetData(sql);
        }
        public bool themMoiGiaChiTieu(Chi_Tieu_Obj obj,bool themmoi)
        {
            string sql = "insert into GIA_CHI_TIEU(MACT,TENCT,GIA,TENPHANLOAIGIA) values ('" + obj.Mact + "',N'" + obj.Tenct + "'," + obj.Gia + ",N'" + obj.Loaigia + "')";
            if (themmoi == false)
                sql = "update GIA_CHI_TIEU set GIA=" + obj.Gia + " where MACT='" + obj.Mact + "' and TENPHANLOAIGIA=N'" + obj.Loaigia + "'";
            return ld.AddData(sql);

        }

        // LAY ALL CHI TIEU  (JOIN GIA_CHI_TIEU)
        public DataTable GetChiTieu()
        {
            string sql = "select * from chi_tieu ct left join gia_chi_tieu gia on ct.mact = gia.mact";
            return ld.GetData(sql);
        }



        // THEM CHI TIEU
        public bool ThemChiTieu(Chi_Tieu_Obj Obj)
        {
            string sql = "insert into  chi_tieu(NHOMCHITIEU,mact,tenct,manenmau,tennenmau,mapp,tenpp,maqc,tenqc,madv,tendv,lod,phan_loai) values (N'"+Obj.Nhom_ct+"','" + Obj.Mact + "', N'" + Obj.Tenct + "','" + Obj.Manenmau + "',N'" + Obj.Tennenmau + "','" + Obj.Mapp + "',N'" + Obj.Tenpp + "','" + Obj.Maqc + "',N'" + Obj.Tenqc + "','" + Obj.Madv + "',N'" + Obj.Tendv + "','" + Obj.Lod + "',N'" + Obj.Phan_loai + "')";
            return ld.AddData(sql);
        }

        // SUA CHI ITEU
        public bool UpdateChiTieu(Chi_Tieu_Obj Obj)
        {
            string sql = "update chi_tieu set NHOMCHITIEU=N'"+Obj.Nhom_ct+"', tenct=N'" + Obj.Tenct + "', manenmau=" + Obj.Manenmau + ", tennenmau=N'" + Obj.Tennenmau + "', mapp=" + Obj.Mapp + ", tenpp=N'" + Obj.Tenpp + "', maqc=" + Obj.Maqc + ", tenqc=N'" + Obj.Tenqc + "', madv='" + Obj.Madv + "', tendv=N'" + Obj.Tendv + "', lod='" + Obj.Lod + "', phan_loai=N'" + Obj.Phan_loai + "' where mact='" + Obj.Mact + "'";
            return ld.AddData(sql);
        }

        //XOA CHITIEU
        public bool XoaChiTieu(Chi_Tieu_Obj Obj,string magia)
        {
           
           
                string sql = "delete GIA_CHI_TIEU where MAGIA=" + magia;
                
                DataTable dt2 = ld.GetData("select * from GIA_CHI_TIEU where MACT='" + Obj.Mact + "'");
                
                if(dt2.Rows.Count<=1)
                 sql = "delete CHI_TIEU where mact='" + Obj.Mact + "'";
                return ld.AddData(sql);
            
        }
        public int kiemtratontaict(Chi_Tieu_Obj Obj)        
        {
            return ld.GetData("select * from PHIEU_YCPT where MACT='" + Obj.Mact + "'").Rows.Count;
        }
        public int kiemtrasoluonggia(Chi_Tieu_Obj Obj)
        {
            return ld.GetData("select * from GIA_CHI_TIEU where MACT='" + Obj.Mact + "'").Rows.Count;
        }

        // SUA GIA CHI TIEU
        public bool UpdateGiaChiTieu(Chi_Tieu_Obj Obj)
        {
            string sql = "";
            return ld.AddData(sql);
        }


        // LAY ALL NEN_MAU, PP, QC, DV, PL GIA
        public DataTable GetNenMau()
        {
            string sql = "select * from nen_mau";
            return ld.GetData(sql);
        }
        public DataTable GetPhuongPhap()
        {
            string sql = "select * from phuong_phap";
            return ld.GetData(sql);
        }
        public DataTable GetQuyChuan()
        {
            string sql = "select * from quy_chuan";
            return ld.GetData(sql);
        }
        public DataTable GetDonVi()
        {
            string sql = "select * from don_vi";
            return ld.GetData(sql);
        }
        public DataTable GetPhanLoaiGia()
        {
            string sql = "select * from phan_loai_gia";
            return ld.GetData(sql);
        }


        // THEM NENMAU
        public bool ThemNenMau(Nen_Mau_Obj Obj)
        {
            string sql = "insert  into nen_mau(tennenmau) values (N'"+Obj.Tennenmau+"')";
            return ld.AddData(sql);
        }


        // THEM phuong_phap
        public bool ThemPhuongPhap(Phuong_Phap_Obj Obj)
        {
            string sql = "insert  into phuong_phap(tenpp) values (N'" + Obj.Tenpp + "')";
            return ld.AddData(sql);
        }


        // THEM QUY_CHUAN
        public bool ThemQuyChuan(Quy_Chuan_Obj Obj)
        {
            string sql = "insert  into quy_chuan(tenqc) values (N'" + Obj.Tenqc + "')";
            return ld.AddData(sql);
        }


        // THEM DON_VI
        public bool ThemDonVi(Don_Vi_Obj Obj)
        {
            string sql = "insert  into don_vi(tendv) values (N'" + Obj.Tendv + "')";
            return ld.AddData(sql);
        }


        // THEM DON_VI
        public bool ThemPhanLoaiGia(Phan_Loai_Gia_Obj Obj)
        {
            string sql = "insert  into phan_loai_gia(tenphanloaigia) values (N'" + Obj.Tenphanloaigia + "')";
            return ld.AddData(sql);
        }


        // kiem tra ton tai NEN_MAU
        public int KTNenMau(Nen_Mau_Obj Obj)
        {
            string sql = "SELECT * FROM NEN_MAU WHERE TENNENMAU LIKE N'"+Obj.Tennenmau+"'";
            DataTable dt = ld.GetData(sql);
            int  count= dt.Rows.Count;
            return count;
        }
        // kiem tra ton tai PHUONG_PHAP
        public int KTPhuongPhap(Phuong_Phap_Obj Obj)
        {
            string sql = "SELECT * FROM PHUONG_PHAP WHERE TENPP LIKE N'" + Obj.Tenpp + "'";
            DataTable dt = ld.GetData(sql);
            int count = dt.Rows.Count;
            return count;
        }
        // kiem tra ton tai QUY_CHUAN
        public int KTQuyChuan(Quy_Chuan_Obj Obj)
        {
            string sql = "SELECT * FROM QUY_CHUAN WHERE TENQC LIKE N'" + Obj.Tenqc + "'";
            DataTable dt = ld.GetData(sql);
            int count = dt.Rows.Count;
            return count;
        }
        // kiem tra ton tai DON_VI
        public int KTDonVi(Don_Vi_Obj Obj)
        {
            string sql = "SELECT * FROM DON_VI WHERE TENDV LIKE N'" + Obj.Tendv + "'";
            DataTable dt = ld.GetData(sql);
            int count = dt.Rows.Count;
            return count;
        }
        // kiem tra ton tai PHAN_LOAI_GIA
        public int KTPhanLoaiGia(Phan_Loai_Gia_Obj Obj)
        {
            string sql = "SELECT * FROM PHAN_LOAI_GIA WHERE TENPHANLOAIGIA = N'" + Obj.Tenphanloaigia + "'";
            DataTable dt = ld.GetData(sql);
            int count = dt.Rows.Count;
            return count;
        }

    }
}