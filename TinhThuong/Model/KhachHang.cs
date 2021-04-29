using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TinhThuong.Model
{
    public class KhachHang
    {
        public string MaKH;
        public string NganhHang;
        public string TenNPP;
        public long ChiTieu;
        public long TienDo1;
        public long TienDo2;
        public long TongDanhSo;

        public double TienDo1Per;
        public double TienDo2Per;
        public double TongDanhSoPer;

        public long ThuongKhac;
        public long ThuongTienDo1;
        public long ThuongTienDo2;
        public long ThuongThang;
        public KhachHang()
        {
            MaKH = string.Empty;
            NganhHang = string.Empty;
            TenNPP = string.Empty;
            ChiTieu = 0;
            TienDo1 = 0;
            TienDo2 = 0;
            TongDanhSo = 0;

            TienDo1Per = 0;
            TienDo2Per = 0;
            TongDanhSoPer = 0;

            ThuongKhac = 0;
            ThuongTienDo1 = 0;
            ThuongTienDo2 = 0;
            ThuongThang = 0;
        }
    }
}
