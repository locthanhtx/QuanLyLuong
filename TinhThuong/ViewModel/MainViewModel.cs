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
        private static MainViewModel _instance = null;
        private SettingModel _settingModel;

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
            _settingModel = new SettingModel();
            _settingModel.LoadSetting();
        }
        public string PathFileKinhDoanh
        {
            get
            {
                return _settingModel.FileKinhDoanh;
            }
            set
            {
                _settingModel.FileKinhDoanh = value;
                OnPropertyChanged(() => PathFileKinhDoanh);
            }
        }

        public string PathFileThuongNPP
        {
            get
            {
                return _settingModel.FileThuong;
            }
            set
            {
                _settingModel.FileThuong = value;
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
                ProcessFileExcel(_settingModel.FileKinhDoanh);
                FillDataToFileNPP(_settingModel.FileThuong);
                _settingModel.SaveSetting();
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
            string log = string.Empty;
            try
            {
                FileInfo existingFile = new FileInfo(fileName);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    //tinh toan
                    foreach (var khachHang in _khachHang)
                    {
                        if (khachHang.Value.ChiTieu > 0)
                        {
                            khachHang.Value.TienDo1Per = ((double)khachHang.Value.TienDo1 / (double)khachHang.Value.ChiTieu);
                            khachHang.Value.TienDo2Per = ((double)khachHang.Value.TienDo2 / (double)khachHang.Value.ChiTieu);
                            khachHang.Value.TongDanhSoPer = ((double)khachHang.Value.TongDanhSo / (double)khachHang.Value.ChiTieu);
                        }
                    }

                    // Sheet ANY Structure Overiew
                    int sheet = 1;
                    string colMKH = _settingModel.NPP_ColMaKH;
                    string colNganh = _settingModel.NPP_ColNganh;
                    string colChiTieu = _settingModel.NPP_ColChiTieu;
                    string colTienDo1 = _settingModel.NPP_ColTienDo1;
                    string colTienDo2 = _settingModel.NPP_ColTienDo2;
                    string colDanhSo = _settingModel.NPP_ColDanhSo;

                    string colTienDo1_PhanTram = _settingModel.NPP_ColTienDo1_PhanTram;
                    string colTienDo2_PhanTram = _settingModel.NPP_ColTienDo2_PhanTram;
                    string colDanhSo_PhanTram = _settingModel.NPP_ColDanhSo_PhanTram;

                    string colThuongKhac = _settingModel.NPP_ColThuongKhac;
                    string colThuongTienDo1 = _settingModel.NPP_ColThuongTienDo1;
                    string colThuongTienDo2 = _settingModel.NPP_ColThuongTienDo2;
                    string colThuongThang = _settingModel.NPP_ColThuongThang;

                    string colCT_ChietKhau = _settingModel.NPP_ColCT_ChietKhau;
                    string colCT_ThuongKhac = _settingModel.NPP_ColCT_ThuongKhac;
                    string colCT_ThuongTienDo = _settingModel.NPP_ColCT_ThuongTienDo;
                    string colCT_ThuongThang = _settingModel.NPP_ColCT_ThuongThang;

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];

                    int rows = worksheet.Dimension.End.Row;
                    string maKH;
                    string nganhHang;
                    double chietKhau = 0;
                    double thuongKhac = 0;
                    double thuongTienDo = 0;
                    double thuongThag = 0;
                    for (int row = 3; row <= rows; row++)
                    {
                        string temp_string = string.Empty;
                        maKH = (worksheet.Cells[colMKH + row.ToString()].Value ?? string.Empty).ToString().Trim();
                        nganhHang = (worksheet.Cells[colNganh + row.ToString()].Value ?? string.Empty).ToString().Trim();

                        if (maKH == "x" || maKH == "X") continue;
                        temp_string = (worksheet.Cells[colCT_ChietKhau + row.ToString()].Value ?? string.Empty).ToString().Trim();
                        chietKhau = string.IsNullOrEmpty(temp_string) ? 0 : double.Parse(temp_string);
                        temp_string = (worksheet.Cells[colCT_ThuongKhac + row.ToString()].Value ?? string.Empty).ToString().Trim();
                        thuongKhac = string.IsNullOrEmpty(temp_string) ? 0 : double.Parse(temp_string);
                        temp_string = (worksheet.Cells[colCT_ThuongTienDo + row.ToString()].Value ?? string.Empty).ToString().Trim();
                        thuongTienDo = string.IsNullOrEmpty(temp_string) ? 0 : double.Parse(temp_string);
                        temp_string = (worksheet.Cells[colCT_ThuongThang + row.ToString()].Value ?? string.Empty).ToString().Trim();
                        thuongThag = string.IsNullOrEmpty(temp_string) ? 0 : double.Parse(temp_string);

                        string key = maKH + nganhHang;
                        key = key.ToLower();
                        log = key;
                        if (_khachHang.ContainsKey(key))
                        {
                            KhachHang temp_KhachHang = _khachHang[key];
                            worksheet.Cells[colChiTieu + row.ToString()].Value = temp_KhachHang.ChiTieu;
                            worksheet.Cells[colTienDo1 + row.ToString()].Value = temp_KhachHang.TienDo1;
                            worksheet.Cells[colTienDo2 + row.ToString()].Value = temp_KhachHang.TienDo2;
                            worksheet.Cells[colDanhSo + row.ToString()].Value = temp_KhachHang.TongDanhSo;

                            worksheet.Cells[colTienDo1_PhanTram + row.ToString()].Value = temp_KhachHang.TienDo1Per;
                            worksheet.Cells[colTienDo2_PhanTram + row.ToString()].Value = temp_KhachHang.TienDo2Per;
                            worksheet.Cells[colDanhSo_PhanTram + row.ToString()].Value = temp_KhachHang.TongDanhSoPer;

                            if (temp_KhachHang.TienDo1Per >= (double)(0.4))
                            {
                                temp_KhachHang.ThuongTienDo1 = (long)(temp_KhachHang.TienDo1 * thuongTienDo * (1 - chietKhau));
                            }
                            if (temp_KhachHang.TienDo2Per >= (double)(0.6) || (temp_KhachHang.TienDo1Per + temp_KhachHang.TienDo2Per) >= (double)1.0)
                            {
                                temp_KhachHang.ThuongTienDo2 = (long)(temp_KhachHang.TienDo2 * thuongTienDo * (1 - chietKhau));
                            }
                            if (temp_KhachHang.TongDanhSoPer >= (double)1.0)
                            {
                                temp_KhachHang.ThuongThang = (long)(temp_KhachHang.TongDanhSo * thuongThag * (1 - chietKhau));
                            }
                            temp_KhachHang.ThuongKhac = (long)(temp_KhachHang.TongDanhSo * thuongKhac);


                            worksheet.Cells[colThuongKhac + row.ToString()].Value = temp_KhachHang.ThuongKhac;
                            worksheet.Cells[colThuongTienDo1 + row.ToString()].Value = temp_KhachHang.ThuongTienDo1;
                            worksheet.Cells[colThuongTienDo2 + row.ToString()].Value = temp_KhachHang.ThuongTienDo2;
                            worksheet.Cells[colThuongThang + row.ToString()].Value = temp_KhachHang.ThuongThang;
                        }
                    }

                    package.Save();
                }
            }
            catch
            {
                throw new NullReferenceException(log);
            }
        }
        public string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        // (A = 1, B = 2...AA = 27...AAA = 703...)
        public int GetColNumberFromName(string columnName)
        {
            char[] characters = columnName.ToUpperInvariant().ToCharArray();
            int sum = 0;
            for (int i = 0; i < characters.Length; i++)
            {
                sum *= 26;
                sum += (characters[i] - 'A' + 1);
            }
            return sum;
        }
        private void ProcessFileExcel(string fileName)
        {
            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                // Sheet ANY Structure Overiew
                int sheetTienDo = int.Parse(_settingModel.SheetTienDo);
                string colMKH = _settingModel.KD_ColMaKH;
                string colNganh = _settingModel.KD_ColNganh;
                string colChiTieu = _settingModel.KD_ColChiTieu;
                string colThucHien = _settingModel.KD_ColThucHien;
                string colTenNPP = _settingModel.KD_ColTenNPP;
                int colNgayStart = GetColNumberFromName(_settingModel.KD_NgayStart);

                int tienDo1From = int.Parse(_settingModel.TienDo1From);
                int tienDo1To = int.Parse(_settingModel.TienDo1To);
                int tienDo2From = int.Parse(_settingModel.TienDo2From);
                int tienDo2To = int.Parse(_settingModel.TienDo2To);

                ExcelWorksheet worksheetTienDo = package.Workbook.Worksheets[sheetTienDo];

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

                    for (int i = tienDo1From - 1; i < tienDo1To; ++i)
                    {
                        string tienDo = (worksheetTienDo.Cells[row, colNgayStart + i].Value ?? string.Empty).ToString().Trim();
                        if (!string.IsNullOrEmpty(tienDo))
                        {
                            khachHang.TienDo1 += Convert.ToInt64(tienDo);
                        }
                    }

                    for (int i = tienDo2From - 1; i < tienDo2To; ++i)
                    {
                        string tienDo = (worksheetTienDo.Cells[row, colNgayStart + i].Value ?? string.Empty).ToString().Trim();
                        if (!string.IsNullOrEmpty(tienDo))
                        {
                            khachHang.TienDo2 += Convert.ToInt64(tienDo);
                        }
                    }

                    string key = khachHang.MaKH + khachHang.NganhHang;

                    key = key.ToLower();

                    if (_khachHang.ContainsKey(key))
                    {
                        KhachHang temp_KhachHang = _khachHang[key];
                        temp_KhachHang.ChiTieu += khachHang.ChiTieu;
                        temp_KhachHang.TienDo1 += khachHang.TienDo1;
                        temp_KhachHang.TienDo2 += khachHang.TienDo2;
                        temp_KhachHang.TongDanhSo += khachHang.TongDanhSo;
                    }
                    else
                    {
                        _khachHang.Add(key, khachHang);
                    }
                }
            }
        }
    }
}
