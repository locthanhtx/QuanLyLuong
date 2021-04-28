using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using TinhThuong.Model;

namespace TinhThuong.ViewModel
{
    class MainViewModel : ViewModelBase
    {
        private Dictionary<string, KhachHang> _khachHang;
        private string _pathFileKinhDoanh;
        private string _pathFileThuongNPP;
        private static MainViewModel _instance = null;
   
        public static MainViewModel Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new MainViewModel();
                }

                return _instance;
            }
        }

        private MainViewModel()
        {
            _instance = this;
            _khachHang = new Dictionary<string, KhachHang>();
        }
        public string PathFileKinhDoanh
        {
            get
            {
                return _pathFileKinhDoanh;
            }
            set
            {
                _pathFileKinhDoanh = value;
                OnPropertyChanged(() => PathFileKinhDoanh);
            }
        }

        public string PathFileThuongNPP
        {
            get
            {
                return _pathFileThuongNPP;
            }
            set
            {
                _pathFileThuongNPP = value;
                OnPropertyChanged(() => PathFileThuongNPP);
            }
        }

        public ICommand OpenFileKDCommand
        {
            get
            {
                return new RelayCommand(OpenFileKDCommandExecute);
            }
        }
        private void OpenFileKDCommandExecute()
        {
            PathFileKinhDoanh = OpenFileCommandExecute();
        }

        public ICommand OpenFileThuongCommand
        {
            get
            {
                return new RelayCommand(OpenFileThuongCommandExecute);
            }
        }
        private void OpenFileThuongCommandExecute()
        {
            PathFileThuongNPP = OpenFileCommandExecute();
        }

        public ICommand ProcessCommand
        {
            get
            {
                return new RelayCommand(ProcessCommandExecute);
            }
        }
        private void ProcessCommandExecute()
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            try
            {
                _khachHang.Clear();
                ProcessFileExcel(_pathFileKinhDoanh);
                FillDataToFileNPP(_pathFileThuongNPP);
            }
            catch (Exception ex)
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Lỗi Không Thể Mở File !!!!" + ex.Message, "Nhắc Nhở Bạn !!!");
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
            Mouse.OverrideCursor = null;
        }

        private string OpenFileCommandExecute()
        {
            string pathFile = string.Empty;
            // Stream myStream = null;
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Mở File Excel";
            theDialog.Filter = "Excel Files(.xlsx)|*.xlsx|Excel Files(*.xlsm) |*.xlsm|Excel Files(.xls)|*.xls";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    pathFile = theDialog.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi Không Thể Mở File !!!!" + ex.Message, "Nhắc Nhở Bạn !!!");
                }
            }
            return pathFile;
        }
        private void FillDataToFileNPP(string fileName)
        {
            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                // Sheet ANY Structure Overiew
                int sheet = 1;
                string colMKH = "A";
                string colNganh = "E";
                string colChiTieu = "S";
                string colDanhSo = "V";


                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];

                int rows = worksheet.Dimension.End.Row;
                string maKH;
                string nganhHang;
                for (int row = 3; row <= rows; row++)
                {
                    maKH = (worksheet.Cells[colMKH + row.ToString()].Value ?? string.Empty).ToString().Trim();
                    nganhHang = (worksheet.Cells[colNganh + row.ToString()].Value ?? string.Empty).ToString().Trim();

                    string key = maKH + nganhHang;

                    if (_khachHang.ContainsKey(key))
                    {
                        KhachHang temp_KhachHang = _khachHang[key];
                        worksheet.Cells[colChiTieu + row.ToString()].Value = temp_KhachHang.ChiTieu;
                        worksheet.Cells[colDanhSo + row.ToString()].Value = temp_KhachHang.TongDanhSo;
                    }
                }

                package.Save();
            }
        }
        private void ProcessFileExcel(string fileName)
        {
            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                // Sheet ANY Structure Overiew
                int sheetTienDo = 3;
                int sheetDataTho = 3;
                string colMKH = "G";
                string colNganh = "F";
                string colChiTieu = "M";
                string colThucHien = "N";
                string colTenNPP = "H";

                ExcelWorksheet worksheetTienDo = package.Workbook.Worksheets[sheetTienDo];
                ExcelWorksheet worksheetDataTho = package.Workbook.Worksheets[sheetTienDo];

                int rows = worksheetTienDo.Dimension.End.Row;
                string chiTieu;
                string thucHien;
                for (int row = 6; row <= rows; row++)
                {
                    KhachHang khachHang = new KhachHang();
                    khachHang.MaKH = (worksheetTienDo.Cells[colMKH + row.ToString()].Value ?? string.Empty).ToString().Trim();
                    khachHang.NganhHang = (worksheetTienDo.Cells[colNganh + row.ToString()].Value ?? string.Empty).ToString().Trim();
                    khachHang.TenNPP = (worksheetTienDo.Cells[colTenNPP + row.ToString()].Value ?? string.Empty).ToString().Trim();
                    chiTieu = (worksheetTienDo.Cells[colChiTieu + row.ToString()].Value ?? string.Empty).ToString().Trim();
                    thucHien = (worksheetTienDo.Cells[colThucHien + row.ToString()].Value ?? string.Empty).ToString().Trim();

                    if (!string.IsNullOrEmpty(chiTieu))
                    {
                        khachHang.ChiTieu = Convert.ToInt64(chiTieu);
                    }
                    if (!string.IsNullOrEmpty(thucHien))
                    {
                        khachHang.TongDanhSo = Convert.ToInt64(thucHien);
                    }


                    string key = khachHang.MaKH + khachHang.NganhHang;

                    if (_khachHang.ContainsKey(key))
                    {
                        KhachHang temp_KhachHang = _khachHang[key];
                        temp_KhachHang.ChiTieu += khachHang.ChiTieu;
                        temp_KhachHang.TongDanhSo += khachHang.TongDanhSo;
                    }
                    else
                    {
                        _khachHang.Add(khachHang.MaKH + khachHang.NganhHang, khachHang);
                    }
                }
            }
        }
    }
}
