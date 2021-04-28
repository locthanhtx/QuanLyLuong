using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace QuanLyLuong
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtPathFile.Text = openFileDialog.FileName;
            }
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(txtPathFile.Text))
            {
                MessageBox.Show("Vui Lòng Chọn File Muốn Xử Lý!!!!");
                return;
            }

        }
    }
}

