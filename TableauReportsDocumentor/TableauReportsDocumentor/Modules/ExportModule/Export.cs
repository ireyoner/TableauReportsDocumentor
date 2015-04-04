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
using TableauReportsDocumentor.ReportDocumet;

namespace TableauReportsDocumentor.Modules.ExportModule
{
    class Export
    {
        public Dictionary<String, Tuple<ExportInterface, int>> exporters { get; private set; }
        private readonly String fallbackExtension = "trd";
        private SaveFileDialog saveFileDialog;

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
            saveFileDialog.Title = "Save prepared aport documentation";
            saveFileDialog.Filter = "Tableau Report Documentator (*.trd)|*.trd|All files (*.*)|*.*";

            this.exporters = new Dictionary<String, Tuple<ExportInterface, int>>();
            exporters.Add(fallbackExtension, new Tuple<ExportInterface, int>(new TrdExport(), 1));
            loadExportres();
        }

        private int saveExporterIndex = 3;
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
            Console.Out.WriteLine(new DocxExport());
            registerExporter(new DocxExport());
        }

        public void registerExporter(ExportInterface exporter)
        {
            if (exporter.menuItemText != null)
            {
                String id = "exporter_" + exporters.Count;
                if (exporter.fileExtinsion != null)
                {
                    if (!exporters.ContainsKey(exporter.fileExtinsion))
                    {
                        id = exporter.fileExtinsion;
                        exporters.Add(id, getExporterTouple(exporter,true));
                        if (saveFileDialog.Filter != null)
                            saveFileDialog.Filter = saveFileDialog.Filter + "|" + exporter.exportSaveFileDialogFilter;
                        else
                            saveFileDialog.Filter = exporter.exportSaveFileDialogFilter;
                    }
                    else
                    {
                        throw new Exception("Exporter for " + exporter.fileExtinsion + " already exists!");
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
            if (exporter.toolBarButtonContent != null)
            {
                var button = new Button();
                button.Click += buttonClick;
                button.Name = "B_" + id;
                button.Content = exporter.toolBarButtonContent;
                exportToolBar.Items.Add(button);
            }
        }

        private void setupExporterMenuItem(String id, ExportInterface exporter)
        {
            var menuItem = new MenuItem();
            menuItem.Click += menuItemClick;
            menuItem.Name = "M_" + id;
            menuItem.Header = exporter.menuItemText;
            if (exporter.menuItemIcone != null)
            {
                menuItem.Icon = exporter.menuItemIcone;
            }
            exportMenu.Items.Add(menuItem);
        }

        public bool Export_Click(object sender, RoutedEventArgs e, ReportDocument document)
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
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(document.fileName);
                saveFileDialog.DefaultExt = exporter.fileExtinsion;
                saveFileDialog.FilterIndex = saveIndex;

                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    var fileExtinsion = Path.GetExtension(saveFileDialog.FileName);
                    if (exporters.ContainsKey(fileExtinsion))
                    {
                        exporter = exporters[fileExtinsion].Item1;
                        return exporter.export(saveFileDialog.FileName, document.xml);

                    }
                    else if (exporters.ContainsKey(fallbackExtension))
                    {
                        exporter = exporters[fallbackExtension].Item1;
                        var OK = exporter.export(saveFileDialog.FileName, document.xml);
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
                return exporter.export(null, document.xml);
            }
        }
    }
}
