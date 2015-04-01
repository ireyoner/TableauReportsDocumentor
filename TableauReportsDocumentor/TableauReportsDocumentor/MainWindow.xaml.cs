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

        public MainWindow()
        {
            InitializeComponent();
            ExportInterface docx = new DocxExport();
            var button = new Button();
            button.Click += ExportButton_Click;
            button.Content = docx.exportFormat;
            //button.Content = docx.icone;
            ExportTB.Items.Add(button);

            MenuItem mi = new MenuItem();
            mi.Click += ExportButton_Click;
            mi.Header = docx.exportFormat;
            ExportM.Items.Add(mi);
        }

        private ReportDocument document = new ReportDocument();


        private void WriteDocument()
        {
            MemoryStream mStream = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(mStream, Encoding.Unicode);
            writer.Formatting = Formatting.Indented;
            writer.Indentation = 4;
            writer.QuoteChar = '\'';
            document.xml.WriteContentTo(writer);
            writer.Flush();
            mStream.Flush();

            // Have to rewind the MemoryStream in order to read
            // its contents.
            mStream.Position = 0;

            // Read MemoryStream contents into a StreamReader.
            StreamReader sReader = new StreamReader(mStream);

            // Extract the text from the StreamReader.
            String FormattedXML = sReader.ReadToEnd();

            outputTest.Text = FormattedXML;
            Console.Write(FormattedXML);

        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            var exp = new Export();
            if (!exp.ExportTableauReportDocumentatnion(document.xml))
            {
                MessageBox.Show("There was an error saving your documentation", "Export error");
            }

        }

        private void Save(object sender, RoutedEventArgs e)
        {
            document.Save();
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            if(document.Open())
                WriteDocument();
        }
    }
}
