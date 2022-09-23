using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMauKiemNghiem.Object
{
    public class Khach_Hang_Obj
    {
        private string makh, tenkh, diachi, masothue, sdt, email, sofax, nguoi_lh, sdt_lh, email_lh;

        public string Makh
        {
            get { return makh; }
            set { makh = value; }
        }

        public string Email_lh
        {
            get { return email_lh; }
            set { email_lh = value; }
        }

        public string Sdt_lh
        {
            get { return sdt_lh; }
            set { sdt_lh = value; }
        }

        public string Nguoi_lh
        {
            get { return nguoi_lh; }
            set { nguoi_lh = value; }
        }

        public string Sofax
        {
            get { return sofax; }
            set { sofax = value; }
        }

        public string Email
        {
            get { return email; }
            set { email = value; }
        }

        public string Sdt
        {
            get { return sdt; }
            set { sdt = value; }
        }

        public string Masothue
        {
            get { return masothue; }
            set { masothue = value; }
        }

        public string Diachi
        {
            get { return diachi; }
            set { diachi = value; }
        }

        public string Tenkh
        {
            get { return tenkh; }
            set { tenkh = value; }
        }


        public Khach_Hang_Obj()
        {
        }
    }
}
