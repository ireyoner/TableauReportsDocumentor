/*
 * File: Export.cs
 * Class: Export
 * 
 * Class responsible for handling any document export related tasks. It registers all of the exporters, creates all of the appropiate UI and 
 * is responsible for calling their exporting functions
 * 
 */
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
        private SaveFileDialog saveFileDialog;
        private String saveFileDialogFilter;

        private RoutedEventHandler buttonClick;
        private ToolBar exportToolBar;

        private RoutedEventHandler menuItemClick;
        private MenuItem exportMenu;

        /* Function: Export
         * A constructor for the Export class with parameters. Responsible for registering event handlers, creating appropiate UI etc.
         * 
         * Parameters:
         * 
         */
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
            saveFileDialogFilter = null;

            this.exporters = new Dictionary<String, Tuple<ExportInterface, int>>();
            loadExporters();
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

        /*
         * Function: loadExporters
         * Function responsible for loading all of exporters. All exporters must be registered here in order to function.
         */
        private void loadExporters()
        {
            RegisterExporter(new ExportDocX());
            RegisterExporter(new ExportDoc());
            RegisterExporter(new ExportCSV());
        }
        
        /*
         * Function: RegisterExporter
         * Registers a single exporter class for work.
         * 
         * Parameters:
         *  exporter - an exporter to be registered
         */
        public void RegisterExporter(ExportInterface exporter)
        {
            if (exporter.MenuItemText != null)
            {
                String id = "exporter_" + exporters.Count;
                if (exporter.FileExtension != null) // check if the exporter will be saving a file on the hard drive
                {
                    if (!exporters.ContainsKey(exporter.FileExtension)) //check if there is no exporter already defined for this file format
                    {
                        id = exporter.FileExtension;
                        exporters.Add(id, getExporterTouple(exporter,true));
                        if (saveFileDialogFilter != null)
                            saveFileDialogFilter = saveFileDialogFilter + "|" + exporter.ExportSaveFileDialogFilter;
                        else
                            saveFileDialogFilter = exporter.ExportSaveFileDialogFilter;
                    }
                    else
                    {
                        throw new Exception("Exporter for " + exporter.FileExtension + " already exists!");
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

        /*
         * Function: setupExporterButton
         * Function responsible for creating a UI button for the new exporter
         * 
         * Parameters:
         *  id - string identifying the exporter in the dictionary
         *  exporter - an exporter to be registered
         */
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

        /*
         * Function: setupExporterMenuItem
         * Function responsible for creating a UI menu item for the new exporter
         * 
         * Parameters:
         *  id - string identifying the exporter in the dictionary
         *  exporter - an exporter to be registered
         */
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

        /*
         * Function: ExportDocument
         * An event handler called whenever there is an attempt of exporting a file.
         */
        public bool ExportDocument(object sender, RoutedEventArgs e, ReportDocumentManager document)
        {
            ExportInterface exporter;
            int saveIndex = noSaveExporterIndex;
            // Get the info about exporter depening on the UI element that triggered the export.
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
                // Display a save file dialog
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(document.FileName);
                saveFileDialog.Filter = saveFileDialogFilter + "|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = exporter.FileExtension;
                saveFileDialog.FilterIndex = saveIndex;

                if (saveFileDialog.ShowDialog()??false)
                {
                    var fileExtension = Path.GetExtension(saveFileDialog.FileName).Substring(1);
                    if (exporters.ContainsKey(fileExtension))
                    {
                        exporter = exporters[fileExtension].Item1;
                        try {
                            //attempt exporting the file.
                            if (exporter.Export(saveFileDialog.FileName, document.GetExportXml()))
                            {
                                var ExportOK = new OpenExportedFile(saveFileDialog.FileName);
                                ExportOK.ShowDialog();
                            }
                        }
                        catch (Exception e2) {
                            var ExportError = new ExportFileError(e2.Message);
                            ExportError.ShowDialog();
                        }
                        return true;
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
