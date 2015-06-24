using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace TableauReportsDocumentor
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            MainWindow = new MainWindow();
            
            // Opening files associeted with application does not work with related paths in settings!

            //MainWindow = null;
            //if (e.Args.Length == 1) //make sure an argument is passed
            //{
            //    FileInfo file = new FileInfo(e.Args[0]);
            //    if (file.Exists) //make sure it's actually a file
            //    {
            //        MainWindow = new MainWindow(e.Args[0]);
            //    }
            //}
            //if (MainWindow != null)
            //{
            //    MainWindow = new MainWindow();
            //}
        }
    }
}
