using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TinhThuong.Model
{
    public class SettingModel
    {
        public string FileKinhDoanh { get; set; }
        public string FileThuong { get; set; }
        public string TienDo1From { get; set; }
        public string TienDo1To { get; set; }
        public string TienDo2From { get; set; }
        public string TienDo2To { get; set; }
        public string SheetTienDo { get; set; }
        public string KD_ColMaKH { get; set; }
        public string KD_ColNganh { get; set; }
        public string KD_ColChiTieu { get; set; }
        public string KD_ColThucHien { get; set; }
        public string KD_ColTenNPP { get; set; }
        public string KD_NgayStart { get; set; }
        public string NPP_ColMaKH { get; set; }
        public string NPP_ColNganh { get; set; }
        public string NPP_ColChiTieu { get; set; }
        public string NPP_ColTienDo1 { get; set; }
        public string NPP_ColTienDo2 { get; set; }
        public string NPP_ColDanhSo { get; set; }
        public string NPP_ColTienDo1_PhanTram { get; set; }
        public string NPP_ColTienDo2_PhanTram { get; set; }
        public string NPP_ColDanhSo_PhanTram { get; set; }
        public string NPP_ColThuongKhac { get; set; }
        public string NPP_ColThuongTienDo1 { get; set; }
        public string NPP_ColThuongTienDo2 { get; set; }
        public string NPP_ColThuongThang { get; set; }
        public string NPP_ColCT_ChietKhau { get; set; }
        public string NPP_ColCT_ThuongKhac { get; set; }
        public string NPP_ColCT_ThuongTienDo { get; set; }
        public string NPP_ColCT_ThuongThang { get; set; }
        public string NPP_ColCT_ThuongQuy3 { get; set; }
        public string NPP_ColCT_ThuongQuy6 { get; set; }
        public string NPP_ColCT_ThuongNam { get; set; }

        public void LoadSetting()
        {
            FileKinhDoanh = ConfigurationManager.AppSettings["FileKinhDoanh"];
            FileThuong = ConfigurationManager.AppSettings["FileThuong"];
            TienDo1From = ConfigurationManager.AppSettings["TienDo1From"];
            TienDo1To = ConfigurationManager.AppSettings["TienDo1To"];
            TienDo2From = ConfigurationManager.AppSettings["TienDo2From"];
            TienDo2To = ConfigurationManager.AppSettings["TienDo2To"];
            SheetTienDo = ConfigurationManager.AppSettings["SheetTienDo"];
            KD_ColMaKH = ConfigurationManager.AppSettings["KD_ColMaKH"];
            KD_ColNganh = ConfigurationManager.AppSettings["KD_ColNganh"];
            KD_ColChiTieu = ConfigurationManager.AppSettings["KD_ColChiTieu"];
            KD_ColThucHien = ConfigurationManager.AppSettings["KD_ColThucHien"];
            KD_ColTenNPP = ConfigurationManager.AppSettings["KD_ColTenNPP"];
            KD_NgayStart = ConfigurationManager.AppSettings["KD_NgayStart"];
            NPP_ColMaKH = ConfigurationManager.AppSettings["NPP_ColMaKH"];
            NPP_ColNganh = ConfigurationManager.AppSettings["NPP_ColNganh"];
            NPP_ColChiTieu = ConfigurationManager.AppSettings["NPP_ColChiTieu"];
            NPP_ColTienDo1 = ConfigurationManager.AppSettings["NPP_ColTienDo1"];
            NPP_ColTienDo2 = ConfigurationManager.AppSettings["NPP_ColTienDo2"];
            NPP_ColDanhSo = ConfigurationManager.AppSettings["NPP_ColDanhSo"];
            NPP_ColTienDo1_PhanTram = ConfigurationManager.AppSettings["NPP_ColTienDo1_PhanTram"];
            NPP_ColTienDo2_PhanTram = ConfigurationManager.AppSettings["NPP_ColTienDo2_PhanTram"];
            NPP_ColDanhSo_PhanTram = ConfigurationManager.AppSettings["NPP_ColDanhSo_PhanTram"];

            NPP_ColThuongKhac = ConfigurationManager.AppSettings["NPP_ColThuongKhac"];
            NPP_ColThuongTienDo1 = ConfigurationManager.AppSettings["NPP_ColThuongTienDo1"];
            NPP_ColThuongTienDo2 = ConfigurationManager.AppSettings["NPP_ColThuongTienDo2"];
            NPP_ColThuongThang = ConfigurationManager.AppSettings["NPP_ColThuongThang"];

            NPP_ColCT_ChietKhau = ConfigurationManager.AppSettings["NPP_ColCT_ChietKhau"];
            NPP_ColCT_ThuongKhac = ConfigurationManager.AppSettings["NPP_ColCT_ThuongKhac"];
            NPP_ColCT_ThuongTienDo = ConfigurationManager.AppSettings["NPP_ColCT_ThuongTienDo"];
            NPP_ColCT_ThuongThang = ConfigurationManager.AppSettings["NPP_ColCT_ThuongThang"];
            NPP_ColCT_ThuongQuy3 = ConfigurationManager.AppSettings["NPP_ColCT_ThuongQuy3"];
            NPP_ColCT_ThuongQuy6 = ConfigurationManager.AppSettings["NPP_ColCT_ThuongQuy6"];
            NPP_ColCT_ThuongNam = ConfigurationManager.AppSettings["NPP_ColCT_ThuongNam"];
        }
        public void SaveSetting()
        {
            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            configuration.AppSettings.Settings["FileKinhDoanh"].Value = FileKinhDoanh;
            configuration.AppSettings.Settings["FileThuong"].Value = FileThuong;
            configuration.AppSettings.Settings["TienDo1From"].Value = TienDo1From;
            configuration.AppSettings.Settings["TienDo1To"].Value = TienDo1To;
            configuration.AppSettings.Settings["TienDo2From"].Value = TienDo2From;
            configuration.AppSettings.Settings["TienDo2To"].Value = TienDo2To;
            configuration.AppSettings.Settings["SheetTienDo"].Value = SheetTienDo;
            configuration.AppSettings.Settings["KD_ColMaKH"].Value = KD_ColMaKH;
            configuration.AppSettings.Settings["KD_ColNganh"].Value = KD_ColNganh;
            configuration.AppSettings.Settings["KD_ColChiTieu"].Value = KD_ColChiTieu;
            configuration.AppSettings.Settings["KD_ColThucHien"].Value = KD_ColThucHien;
            configuration.AppSettings.Settings["KD_ColTenNPP"].Value = KD_ColTenNPP;
            configuration.AppSettings.Settings["KD_NgayStart"].Value = KD_NgayStart;
            configuration.AppSettings.Settings["NPP_ColMaKH"].Value = NPP_ColMaKH;
            configuration.AppSettings.Settings["NPP_ColNganh"].Value = NPP_ColNganh;
            configuration.AppSettings.Settings["NPP_ColChiTieu"].Value = NPP_ColChiTieu;
            configuration.AppSettings.Settings["NPP_ColTienDo1"].Value = NPP_ColTienDo1;
            configuration.AppSettings.Settings["NPP_ColTienDo2"].Value = NPP_ColTienDo2;
            configuration.AppSettings.Settings["NPP_ColDanhSo"].Value = NPP_ColDanhSo;

            configuration.AppSettings.Settings["NPP_ColTienDo1_PhanTram"].Value = NPP_ColTienDo1_PhanTram;
            configuration.AppSettings.Settings["NPP_ColTienDo2_PhanTram"].Value = NPP_ColTienDo2_PhanTram;
            configuration.AppSettings.Settings["NPP_ColDanhSo_PhanTram"].Value = NPP_ColDanhSo_PhanTram;

            configuration.AppSettings.Settings["NPP_ColThuongKhac"].Value = NPP_ColThuongKhac;
            configuration.AppSettings.Settings["NPP_ColThuongTienDo1"].Value = NPP_ColThuongTienDo1;
            configuration.AppSettings.Settings["NPP_ColThuongTienDo2"].Value = NPP_ColThuongTienDo2;
            configuration.AppSettings.Settings["NPP_ColThuongThang"].Value = NPP_ColThuongThang;

            configuration.AppSettings.Settings["NPP_ColCT_ChietKhau"].Value = NPP_ColCT_ChietKhau;
            configuration.AppSettings.Settings["NPP_ColCT_ThuongKhac"].Value = NPP_ColCT_ThuongKhac;
            configuration.AppSettings.Settings["NPP_ColCT_ThuongTienDo"].Value = NPP_ColCT_ThuongTienDo;
            configuration.AppSettings.Settings["NPP_ColCT_ThuongThang"].Value = NPP_ColCT_ThuongThang;
            configuration.AppSettings.Settings["NPP_ColCT_ThuongQuy3"].Value = NPP_ColCT_ThuongQuy3;
            configuration.AppSettings.Settings["NPP_ColCT_ThuongQuy6"].Value = NPP_ColCT_ThuongQuy6;
            configuration.AppSettings.Settings["NPP_ColCT_ThuongNam"].Value = NPP_ColCT_ThuongNam;
        }
    }
}
