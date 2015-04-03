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
            exporters = new Dictionary<string, ExportInterface>();

            loadExportres();
        }

        private void loadExportres()
        {
            // TO DO: tutaj dodać pobieranie dynamiczne zamiast tego
            registerExporter(new DocxExport());
        }

        public void registerExporter(ExportInterface exporter)
        {
            if (exporter.menuItemText != null)
            {
                String id = "exporter_" + exporters.Count;
                exporters.Add(id, exporter);
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
            if (sender.GetType() == typeof(Button))
            {
                var button = (Button)sender;
                exporter = exporters[button.Name.Substring(2)];
            }
            else if (sender.GetType() == typeof(MenuItem))
            {
                var menuItem = (MenuItem)sender;
                exporter = exporters[menuItem.Name.Substring(2)];
            }
            else
            {
                return false;
            }

            if (exporter.exportSaveFileDialogFilter != null)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = exporter.exportSaveFileDialogFilter;
                saveFileDialog.Title = "Save prepared aport documentation";
                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    return exporter.export(saveFileDialog.FileName, document.xml);
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return exporter.export(null, document.xml);
            }
        }
    }
}
