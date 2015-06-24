/*
 * File: Import.cs
 * Class: Import
 * 
 * An class responsible for creating list of importers and their menagement
 * 
 */
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.IO.Compression;
using System.Windows;
using TableauReportsDocumentor.Data;
using System.Windows.Controls;

namespace TableauReportsDocumentor.Modules.ImportModule
{
    class Import
    {
        public Dictionary<String, ImportInterface> importers { get; private set; }
        private RoutedEventHandler menuItemClick;
        private MenuItem importMenu;

        /* Function: Import
         * A constructor for the Import class with parameters. Responsible for registering event handlers, creating appropiate UI etc.
         * 
         * Parameters:
         * 
         */
        public Import(
            RoutedEventHandler menuItemClick,
            MenuItem importMenu)
        {
            this.menuItemClick = menuItemClick;
            this.importMenu = importMenu;

            this.importers = new Dictionary<String, ImportInterface>();
            loadImportres();
        }

        /* Function: loadImportres
         * Creates list of possible importers, currently only one
         * 
         * Parameters:
         * 
         */
        private void loadImportres()
        {
            RegisterImporter(new ImportTWBandTWBX());
            // TO DO: tutaj dodać pobieranie dynamiczne zamiast tego
        }

        /* Function: RegisterImporter
         * Register new Importer
         * 
         * Parameters:
         *  ImportInterface importer - registered Importer
         */
        public void RegisterImporter(ImportInterface importer)
        {
            if (importer.MenuItemText != null)
            {
                String id = "importer_" + importers.Count;
                importers.Add(id, importer);
                setupExporterMenuItem(id, importer);
            }
        }

        /* Function: setupExporterMenuItem
         * Adds new Importer to menu
         * 
         * Parameters:
         *  String id - visible name
         *  ImportInterface importer - Importer to add
         */
        private void setupExporterMenuItem(String id, ImportInterface importer)
        {
            var menuItem = new MenuItem();
            menuItem.Click += menuItemClick;
            menuItem.Name = id;
            menuItem.Header = importer.MenuItemText;
            if (importer.MenuItemIcone != null)
            {
                menuItem.Icon = importer.MenuItemIcone;
            }
            importMenu.Items.Add(menuItem);
        }

        /* Function: ImportDocument
         * Function called after click on menu item
         * 
         * Parameters:
         *  
         * Returns:
         *  ReportDocumentManager - imported document
         */
        public ReportDocumentManager ImportDocument(object sender, RoutedEventArgs e)
        {
            ImportInterface importer;
            if (sender.GetType() == typeof(MenuItem))
            {
                var menuItem = (MenuItem)sender;
                importer = importers[menuItem.Name];
            }
            else
            {
                return null;
            }

            importer.Import();

            var xmlDocument = new ReportContent(importer.OriginalReport,importer.ConvertedReport);
            if (xmlDocument != null)
            {
                return new ReportDocumentManager(xmlDocument);
            }
            return null;

        }
    }
}
