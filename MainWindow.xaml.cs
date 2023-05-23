using Atlas.Data;
using Atlas.Interops.LibreOffice;
using Atlas.Interops.Office;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Atlas
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            string pwd = Environment.CurrentDirectory;
            string prjFolder = pwd.Remove(pwd.IndexOf("Atlas") + 6);
            DocAttributes da = new DocAttributes();
            using (CalcReader cr = new CalcReader())
                da = cr.PullAttributes(prjFolder + @"Reference\10.05.04___2021_56_.plx.xlsx");
        }
    }
}
