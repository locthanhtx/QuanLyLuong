using OfficeOpenXml;
using SendMail.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendMail.ViewModel
{
    class MainViewModel : ViewModelBase
    {
        private string _pathFile;
        private string _content;
        private static MainViewModel _instance = null;
        private List<InfoUser> _userInfo;
        private bool hasCheck;
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
            hasCheck = false;
        }

        public string PathFile
        {
            get
            {
                return _pathFile;
            }
            set
            {
                _pathFile = value;
                OnPropertyChanged(() => PathFile);
            }
        }

        public string Content
        {
            get
            {
                return _content;
            }
            set
            {
                _content = value;
                OnPropertyChanged(() => Content);
            }
        }
        public ICommand OpenFileCommand
        {
            get
            {
                return new RelayCommand(OpenFileCommandExecute);
            }
        }

        public ICommand SentEmailCommand
        {
            get
            {
                return new RelayCommand(SentEmailCommandExecute);
            }
        }

        public ICommand CheckContentCommand
        {
            get
            {
                return new RelayCommand(CheckContentCommandExecute);
            }
        }
        private void SentEmailCommandExecute()
        {
            if (hasCheck == false)
            {
                MessageBox.Show("Vui Lòng Nhập Nội Dung Và Nhấn Kiểm Tra Trước... Pờ Ly", "Nhắc Nhở Bạn !!!");
                return;
            }

            if (string.IsNullOrEmpty(_pathFile))
            {
                MessageBox.Show("Vui Lòng Chọn File Trước... Pờ Ly", "Nhắc Nhở Bạn !!!");
                return;
            }

            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                string sig = ReadSignature();
                Outlook.Application app = new Outlook.Application();

                foreach (var user in _userInfo)
                {
                    Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);

                    if (user.email.Count > 0)
                    {
                        mailItem.To = string.Join(";", user.email.ToArray());
                    }

                    mailItem.Subject = user.subject;
                    user.content = user.content.Replace(System.Environment.NewLine, "<br>");
                    mailItem.Body = user.content + "<br><br>" + sig;

                    foreach (var file in user.fileName)
                    {
                        mailItem.Attachments.Add(file, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }

                    mailItem.HTMLBody = mailItem.Body;
                    mailItem.Send();
                }
                Mouse.OverrideCursor = null;
            }
            catch (Exception sysEx)
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Lỗi Gửi Mail !!!! " + sysEx.Message, "Nhắc Nhở Bạn !!!");
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
            MessageBox.Show("Xong Gồi", "Nhắc Nhở Bạn !!!");
        }

        private void CheckContentCommandExecute()
        {
            if (string.IsNullOrEmpty(_content))
            {
                MessageBox.Show("Vui Lòng Nhập Nội Dung Để Gửi Mail Nha Bạn... Pờ Ly", "Nhắc Nhở Bạn !!!");
                return;
            }

            string temp = string.Empty;

            foreach (var user in _userInfo)
            {
                string temp_content = string.Format(_content, user.hoTen);
                user.content = temp_content;

                temp += "Tiêu Đề: " + user.subject + System.Environment.NewLine + System.Environment.NewLine;
                temp += user.content + System.Environment.NewLine + System.Environment.NewLine;
                temp += "Được Gửi Tới: " + string.Join(";", user.email.ToArray()) + System.Environment.NewLine + System.Environment.NewLine;
                temp += "File Đính Kèm: " + string.Join(";", user.fileName.ToArray()) + System.Environment.NewLine + System.Environment.NewLine;
                temp += "----------------------------------------------------------------" + System.Environment.NewLine;
            }

            _content = temp;
            hasCheck = true;
            OnPropertyChanged(() => Content);
        }
        private void OpenFileCommandExecute()
        {
            // Stream myStream = null;
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Mở File Excel";
            theDialog.Filter = "Excel Files(*.xlsm)|*.xlsm|Excel Files(.xlsx)|*.xlsx|Excel Files(.xls)|*.xls";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ProcessFileExcel(theDialog.FileName);
                    _pathFile = theDialog.FileName;
                    OnPropertyChanged(() => PathFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi Không Thể Mở File !!!!" + ex.Message, "Nhắc Nhở Bạn !!!");
                }
            }
        }

        private string ReadSignature()
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                if (fiSignature.Length > 0)
                {
                    StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                    signature = sr.ReadToEnd();

                    if (!string.IsNullOrEmpty(signature))
                    {
                        string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName + "_files/");
                    }
                }
            }
            return signature;
        }

        private void ProcessFileExcel(string fileName)
        {
            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                _userInfo = new List<InfoUser>();

                // Sheet ANY Structure Overiew
                int sheetInfo = 1;
                int colName = 1;
                int colSubject = 2;
                int colEmail = 3;
                int colFileName = 4;

                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetInfo];

                int rows = worksheet.Dimension.End.Row;

                for (int row = 2; row <= rows; row++)
                {
                    InfoUser user = new InfoUser();
                    user.hoTen = worksheet.Cells[row, colName].Value.ToString().Trim();
                    user.subject = worksheet.Cells[row, colSubject].Value.ToString().Trim();
                    user.email = worksheet.Cells[row, colEmail].Value.ToString().Trim().Split(';').ToList();
                    user.fileName = worksheet.Cells[row, colFileName].Value.ToString().Trim().Split(';').ToList();

                    _userInfo.Add(user);
                }
            }
        }
    }
}
