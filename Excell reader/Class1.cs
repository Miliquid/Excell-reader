using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace Excell_reader
{
    class Class1 : Window
    {
        public void FindFile()
            {
        var d = new OpenFileDialog();



        d.InitialDirectory = @"C:\\";
            d.Filter = "excel files (*.xls, *.xlsx)|*.xls;*.xlsx";
            d.FilterIndex = 2;
            d.RestoreDirectory = true;
        }
    } 
