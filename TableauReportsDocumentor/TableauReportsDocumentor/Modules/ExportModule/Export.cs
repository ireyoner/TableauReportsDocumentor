using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using TableauReportsDocumentor.Export_Converters;
using TableauReportsDocumentor.Data;

namespace TableauReportsDocumentor.Modules.ExportModule
{
    class Export
    {
        public Dictionary<String, Tuple<ExportInterface, int>> exporters { get; private set; }
        private readonly String fallbackExtension = "trd";
        private SaveFileDialog saveFileDialog;
        private String saveFileDialogFilter;

        private RoutedEventHandler buttonClick;
        private ToolBar exportToolBar;

        private RoutedEventHandler menuItemClick;
        private MenuItem exportMenu;

        public Export(
            RoutedEventHandler buttonClick,
            ToolBar exportToolBar,
            RoutedEventHandler menuItemClick,
            MenuItem exportMenu)
        {
            this.buttonClick = buttonClick;
            this.exportToolBar = exportToolBar;
            this.menuItemClick = menuItemClick;
            this.exportMenu = exportMenu;

            this.saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save prepared report documentation";
            saveFileDialogFilter = "Tableau Report Documentator (*.trd)|*.trd";

            this.exporters = new Dictionary<String, Tuple<ExportInterface, int>>();
            exporters.Add(fallbackExtension, getExporterTouple(new TrdExport(),true));
            loadExportres();
        }

        private int saveExporterIndex = 1;
        private const int noSaveExporterIndex = -1;
        private Tuple<ExportInterface, int> getExporterTouple(ExportInterface exporter, bool isForSave)
        {
            if (isForSave)
                return new Tuple<ExportInterface, int>(exporter, saveExporterIndex++);
            else
                return new Tuple<ExportInterface, int>(exporter, noSaveExporterIndex);
        }

        private void loadExportres()
        {
            // TO DO: tutaj dodać pobieranie dynamiczne zamiast tego
            //RegisterExporter(new DocxExport());
            RegisterExporter(new ExportDocX());
            RegisterExporter(new ExportCSV());
        }

        public void RegisterExporter(ExportInterface exporter)
        {
            if (exporter.MenuItemText != null)
            {
                String id = "exporter_" + exporters.Count;
                if (exporter.FileExtinsion != null)
                {
                    if (!exporters.ContainsKey(exporter.FileExtinsion))
                    {
                        id = exporter.FileExtinsion;
                        exporters.Add(id, getExporterTouple(exporter,true));
                        if (saveFileDialogFilter != null)
                            saveFileDialogFilter = saveFileDialogFilter + "|" + exporter.ExportSaveFileDialogFilter;
                        else
                            saveFileDialogFilter = exporter.ExportSaveFileDialogFilter;
                    }
                    else
                    {
                        throw new Exception("Exporter for " + exporter.FileExtinsion + " already exists!");
                    }
                }
                else
                {
                    exporters.Add(id, getExporterTouple(exporter, false));
                }
                setupExporterButton(id, exporter);
                setupExporterMenuItem(id, exporter);
            }
        }

        private void setupExporterButton(String id, ExportInterface exporter)
        {
            if (exporter.ToolBarButtonContent != null)
            {
                var button = new Button();
                button.Click += buttonClick;
                button.Name = "B_" + id;
                button.Content = exporter.ToolBarButtonContent;
                exportToolBar.Items.Add(button);
            }
        }

        private void setupExporterMenuItem(String id, ExportInterface exporter)
        {
            var menuItem = new MenuItem();
            menuItem.Click += menuItemClick;
            menuItem.Name = "M_" + id;
            menuItem.Header = exporter.MenuItemText;
            if (exporter.MenuItemIcone != null)
            {
                menuItem.Icon = exporter.MenuItemIcone;
            }
            exportMenu.Items.Add(menuItem);
        }

        public bool ExportDocument(object sender, RoutedEventArgs e, ReportDocumentMenager document)
        {
            ExportInterface exporter;
            int saveIndex = noSaveExporterIndex;
            if (sender.GetType() == typeof(Button))
            {
                var button = (Button)sender;
                exporter = exporters[button.Name.Substring(2)].Item1;
                saveIndex = exporters[button.Name.Substring(2)].Item2;
            }
            else if (sender.GetType() == typeof(MenuItem))
            {
                var menuItem = (MenuItem)sender;
                exporter = exporters[menuItem.Name.Substring(2)].Item1;
                saveIndex = exporters[menuItem.Name.Substring(2)].Item2;
            }
            else
            {
                return false;
            }

            if (saveIndex != noSaveExporterIndex)
            {
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(document.FileName);
                saveFileDialog.Filter = saveFileDialogFilter + "|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = exporter.FileExtinsion;
                saveFileDialog.FilterIndex = saveIndex;

                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    var fileExtinsion = Path.GetExtension(saveFileDialog.FileName).Substring(1);
                    if (exporters.ContainsKey(fileExtinsion))
                    {
                        exporter = exporters[fileExtinsion].Item1;
                        return exporter.Export(saveFileDialog.FileName, document.GetExportXml());
                    }
                    else if (exporters.ContainsKey(fallbackExtension))
                    {
                        exporter = exporters[fallbackExtension].Item1;
                        var OK = exporter.Export(saveFileDialog.FileName, document.Content.ConvertedXml);
                        if (OK)
                        {
                            document.FullFilePath = saveFileDialog.FileName;
                        }
                        return OK;
                    }
                }
                return false;
            }
            else
            {
                return exporter.Export(null, document.Content.ConvertedXml);
            }
        }
    }
}
