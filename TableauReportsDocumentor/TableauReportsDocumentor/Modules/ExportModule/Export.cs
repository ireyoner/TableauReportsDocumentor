using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
        public Dictionary<String, ExportInterface> exporters { get; private set; }

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

            exporters = loadExportres();
        }

        private Dictionary<string, ExportInterface> loadExportres()
        {
            var exporters = new Dictionary<string, ExportInterface>();

            // TO DO: tutaj dodać pobieranie dynamiczne zamiast tego
            registerExporter(new DocxExport());
            //registerExporter(new DocxExport());

            return exporters;
        }

        public void registerExporter(ExportInterface exporter)
        {
            exporters.Add(exporter.displayName, exporter);
            setupExporterButton(exporter);
            setupExporterMenuItem(exporter);
        }

        private void setupExporterButton(ExportInterface exporter)
        {
            var button = new Button();
            button.Click += buttonClick;
            button.ToolTip = exporter.displayName;

            // TO DO: poprawić ustawianie ikony zamiast/obok teksu
            button.Content = exporter.exportFormat;
            //button.Content = docx.icone;

            exportToolBar.Items.Add(button);
        }

        private void setupExporterMenuItem(ExportInterface exporter)
        {
            var menuItem = new MenuItem();
            menuItem.Click += menuItemClick;
            menuItem.ToolTip += exporter.displayName;

            // TO DO: poprawić ustawianie ikony 
            //button.Content = docx.icone;
            menuItem.Header = exporter.displayName;

            exportMenu.Items.Add(menuItem);
        }

        public bool Export_Click(object sender, RoutedEventArgs e, ReportDocument document)
        {
            ExportInterface exporter;
            if (sender.GetType() == typeof(Button))
            {
                var button = (Button)sender;
                exporter = exporters[(String)button.ToolTip];
            }
            else if (sender.GetType() == typeof(MenuItem))
            {
                var menuItem = (MenuItem)sender;
                exporter = exporters[(String)menuItem.ToolTip];
            }
            else
            {
                return false;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = exporter.exportFormatFilter;
            saveFileDialog.Title = "Save prepared aport documentation";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                return exporter.export(saveFileDialog.FileName, document.xml);

            }
            return false;
        }
    }
}
