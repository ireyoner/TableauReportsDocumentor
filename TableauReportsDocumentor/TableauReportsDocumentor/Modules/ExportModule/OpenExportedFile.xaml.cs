using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Windows.Shapes;

namespace TableauReportsDocumentor.Modules.ExportModule
{
    /// <summary>
    /// Interaction logic for OpenExportedFile.xaml
    /// </summary>
    public partial class OpenExportedFile : Window
    {
        public OpenExportedFile()
        {
            InitializeComponent();
        }

        public OpenExportedFile(String filepath)
        {
            InitializeComponent();
            this.Filepath = filepath;
        }

        public String Filepath
        {
            get { return FilePath.Text; }
            set { FilePath.Text = value; }
        }

        private void OpenFileLocation(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(Filepath);
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Open Failed!");
            }
        }

        private void OpenFileInOtherProgram(object sender, RoutedEventArgs e)
        {
            try
            {
                Process ExplorerWindowProcess = new Process();
                ExplorerWindowProcess.StartInfo.FileName = "explorer.exe";
                ExplorerWindowProcess.StartInfo.Arguments = "/select,\"" + Filepath + "\"";
                ExplorerWindowProcess.Start();
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Open Failed!");
            }
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
