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
using System.Xml;
using TableauReportsDocumentor.Export_Converters;
using TableauReportsDocumentor.Modules.ExportModule;
using TableauReportsDocumentor.Modules.ImportModule;
using TableauReportsDocumentor.ReportDocumet;

namespace TableauReportsDocumentor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        Export exporter;
        private ReportDocument document;

        public MainWindow()
        {
            InitializeComponent();
            exporter = new Export(ExportButton_Click, ExportTB, ExportButton_Click, ExportM);
            document = new ReportDocument();
        }

        private void WriteDocument()
        {
            outputTest.Text = document.getAsString();
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!exporter.Export_Click(sender, e, document))
            {
                MessageBox.Show("There was an error saving your documentation", "Export error");
            }
            else
            {
                MessageBox.Show("Documentation saved.", "Export OK");
            }
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            document.Save();
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            bool ok = false;
            try
            {
                ok = document.Open();
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Open error!");
            }
            if (ok)
                WriteDocument();
        }
    }
}
