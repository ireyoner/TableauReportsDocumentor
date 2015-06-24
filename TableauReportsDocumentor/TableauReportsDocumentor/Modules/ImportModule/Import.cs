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

        public Import(
            RoutedEventHandler menuItemClick,
            MenuItem importMenu)
        {
            this.menuItemClick = menuItemClick;
            this.importMenu = importMenu;

            this.importers = new Dictionary<String, ImportInterface>();
            loadImportres();
        }

        private void loadImportres()
        {
            RegisterImporter(new ImportTWBandTWBX());
            // TO DO: tutaj dodać pobieranie dynamiczne zamiast tego
        }

        public void RegisterImporter(ImportInterface importer)
        {
            if (importer.MenuItemText != null)
            {
                String id = "importer_" + importers.Count;
                importers.Add(id, importer);
                setupExporterMenuItem(id, importer);
            }
        }

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
