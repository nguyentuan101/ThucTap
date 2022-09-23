using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyMauKiemNghiem.Object;
using System.Data;
namespace QuanLyMauKiemNghiem.Model
{
    public class XuatYCPT_Mod
    {
        LoadData ld = new LoadData();

        public DataTable getMauChuaXuat()
        {
            string sql = "select distinct  phieu.MAHD, TENKH, JOBNO, MAMAU, TENMAU, NGAYXUATPHIEU_YCPT, TIME_NHANMAU_TT  from HOP_DONG as hd , PHIEU_YCPT as phieu  WHERE hd.MAHD=phieu.MAHD and NGAYXUATPHIEU_YCPT IS NULL OR NGAYXUATPHIEU_YCPT=''";
            return ld.GetData(sql);
        }
        public DataTable getMauChuaXuat(string maNK)
        {
            string sql = "select distinct  phieu.MAHD, TENKH, JOBNO, MAMAU, TENMAU, NGAYXUATPHIEU_YCPT, TIME_NHANMAU_TT  from HOP_DONG as hd, PHIEU_YCPT as phieu WHERE hd.MAHD=phieu.MAHD and MANK = N'" + maNK + "' and NGAYXUATPHIEU_YCPT IS NULL OR NGAYXUATPHIEU_YCPT='' ";
            return ld.GetData(sql);
        }

        public DataTable getMauDaXuat()
        {
            string sql = "select distinct  phieu.MAHD, TENKH, JOBNO, MAMAU, TENMAU, NGAYXUATPHIEU_YCPT, TIME_NHANMAU_TT  from HOP_DONG as hd, PHIEU_YCPT as phieu WHERE hd.MAHD=phieu.MAHD and NGAYXUATPHIEU_YCPT IS NOT NULL";
            return ld.GetData(sql);
        }
        public DataTable getMauDaXuat(string maNK)
        {
            string sql = "select distinct  phieu.MAHD, TENKH, JOBNO, MAMAU, TENMAU, NGAYXUATPHIEU_YCPT, TIME_NHANMAU_TT  from HOP_DONG as hd, PHIEU_YCPT as phieu WHERE hd.MAHD=phieu.MAHD and MANK = N'" + maNK + "' and NGAYXUATPHIEU_YCPT IS NOT NULL OR NGAYXUATPHIEU_YCPT=''";
            return ld.GetData(sql);
        }

        public DataTable GetNoiKiem()
        {
            string sql = "select * from NOI_KIEM";
            return ld.GetData(sql);
        }












    }
}
