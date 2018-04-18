using System;
using System.Collections.Generic;
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
//using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel;
namespace Excell_reader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Application.Application app;
        private Application.Workbooks wrbks=null;
        public Application.Workbook wrbk = null;
        private Application.Worksheet wrsh;
        public string b;
        
        public MainWindow()
        {
            InitializeComponent();

        }


        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog();


            d.InitialDirectory = @"C:\\";
            d.Filter = "excel files (*.xls, *.xlsx)|*.xls;*.xlsx";
            d.FilterIndex = 2;
            d.RestoreDirectory = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // string kj= Path.GetDirectoryName(d.FileName);
                Sfile.Text = d.FileName;
            }

        }

        private void OpenFileF_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog();


            d.InitialDirectory = @"C:\\";
            d.Filter = "excel files (*.xls, *.xlsx)|*.xls;*.xlsx";
            d.FilterIndex = 2;
            d.RestoreDirectory = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                FFile.Text = d.FileName;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {app = new Application.Application { DisplayAlerts = false };
            try
            {
                wrbks = app.Workbooks;
                wrbk = wrbks.Open(Path.Combine(Environment.CurrentDirectory, FFile.Text));
                wrsh = wrbk.ActiveSheet as Application.Worksheet;
            wrsh.Range["A1"].Value= "Hello World";
            app.Visible = true;
            Topmost = true;
                wrbk.Save();
            }
            finally
            {
                app.ActiveWorkbook.Close();
                app.Quit();

                Marshal.ReleaseComObject(wrsh);
                Marshal.ReleaseComObject(wrbk);
                Marshal.ReleaseComObject(wrbks);
            }
            
            

            
            
        }
    }
}