﻿using Microsoft.Win32;
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
using TableauReportsDocumentor.Modules.EditModule;
using TableauReportsDocumentor.Modules.ImportModule;
using TableauReportsDocumentor.Data;
using ICSharpCode.AvalonEdit.Folding;
using System.Globalization;


namespace TableauReportsDocumentor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private Export exporter;
        private Import importer;
        private ReportDocumentMenager document;

        private Export Exporter { get { return exporter; } }
        private Import Importer { get { return importer; } }
        private ReportDocumentMenager Document { get { return document; } set { document = value; } }

        public MainWindow()
        {
            InitializeComponent();
            EditorView.statusLabel = statusLabel;
            try
            {
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
                Document = new ReportDocumentMenager();
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Init Error!");
                this.Close();
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            var loading = new Loading("Exporting Document");
            loading.Show();

            try
            {
                Exporter.ExportDocument(sender, e, Document);
            }
            catch (Exception e2)
            {
                var ExportError = new ExportFileError(e2.Message);
                ExportError.ShowDialog();
            }
            loading.Close();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var loading = new Loading("Importing Document");
            loading.Show();
            try
            {
                var document = Importer.ImportDocument(sender, e);
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Import error");
            }
            loading.Close();
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            var loading = new Loading("Saving Document");
            loading.Show();
            Document.Save();
            loading.Close();
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            var loading = new Loading("Opening Document");
            loading.Show();
            Boolean status = false;
            try
            {
                status = Document.Open();
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Open error!");
            }
            if (status)
                WriteDocument();
            loading.Close();
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
            EditorView.Report = Document.Content;
        }

        private void SaveAs(object sender, RoutedEventArgs e)
        {
            var loading = new Loading("Saving Document");
            loading.Show();
            Document.SaveAs();
            loading.Close();
        }

        private void Close(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }

    }
}
