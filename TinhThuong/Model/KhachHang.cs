using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TinhThuong.Model
{
    class KhachHang
    {
        public string MaKH;
        public string NganhHang;
        public string TenNPP;
        public long ChiTieu;
        public long TienDo1;
        public long TienDo2;
        public long TongDanhSo;
        public KhachHang()
        {
            MaKH = string.Empty;
            NganhHang = string.Empty;
            TenNPP = string.Empty;
            ChiTieu = 0;
            TienDo1 = 0;
            TienDo2 = 0;
            TongDanhSo = 0;
        }
    }
}
