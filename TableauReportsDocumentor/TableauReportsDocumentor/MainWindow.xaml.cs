using System;
using System.Collections.Generic;
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
using TableauReportsDocumentor.Modules.ExportModule;
using TableauReportsDocumentor.Modules.ImportModule;

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
        }

        private XmlDocument doc;


        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var imp = new Import();
            doc = imp.ImportTableauWorkbooks();
            outputTest.Text = doc.OuterXml;

        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            var exp = new Export();
            if (!exp.ExportTableauReportDocumentatnion(doc))
            {
                MessageBox.Show("There was an error saving your documentation", "Export error");
            }

        }
    }
}
