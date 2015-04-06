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
using TableauReportsDocumentor.AppWindows;
using TableauReportsDocumentor.Export_Converters;
using TableauReportsDocumentor.Modules.ExportModule;
using TableauReportsDocumentor.Modules.ImportModule;
using TableauReportsDocumentor.Data;

namespace TableauReportsDocumentor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private Export exporter;
        private Import importer;
        private ReportDocument document;
        
        private Export Exporter { get { return exporter; } }
        private Import Importer { get { return importer; } }
        private ReportDocument Document { get { return document; } set { document = value; } }

        public MainWindow()
        {
            InitializeComponent();
            importer = new Import(ImportButton_Click, ImportM);
            if (!ImportM.HasItems)
            {
                ImportM.IsEnabled = false;
            }
            exporter = new Export(ExportButton_Click, ExportTB, ExportButton_Click, ExportM);
            if (!ExportM.HasItems)
            {
                ImportM.IsEnabled = false;
            }
            Document = new ReportDocument();
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!Exporter.ExportDocument(sender, e, Document))
            {
                MessageBox.Show("There was an error saving your documentation", "Export error");
            }
            else
            {
                MessageBox.Show("Documentation saved.", "Export OK");
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var document = Importer.ImportDocument(sender, e);
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Import error!");
            }
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            Document.Save();
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            try
            {
                Document.Open();
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Open error!");
            }
            WriteDocument();
        }

        private void SettingsWndShow(object sender, RoutedEventArgs e)
        {
            new SettingsWindow().Show();
        }

        private void AboutWndShow(object sender, RoutedEventArgs e)
        {
            new About().Show();
        }

        private void WriteDocument()
        {
            outputTest.Text = Document.GetAsString();
        }

    }
}
