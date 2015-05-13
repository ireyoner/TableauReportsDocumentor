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
using TableauReportsDocumentor.Data;

namespace TableauReportsDocumentor.Modules.EditModule
{
    /// <summary>
    /// Interaction logic for TreeView.xaml
    /// </summary>
    public partial class ReportTreeView : UserControl
    {
        private XmlDataProvider dataProvider;
        private ReportContent report;
        public ReportContent Report
        {
            get { return report; }
            set
            {
                report = value;
                ReloadTreeView();
            }
        }
        public Label statusLabel { get; set; }

        public ReportTreeView()
        {
            InitializeComponent();
            dataProvider = (XmlDataProvider)this.FindResource("ReportXml");
        }

        private void TreeElementModified(object sender, RoutedEventArgs e)
        {
            try
            {
                Report.ConvertedXml = dataProvider.Document;
                statusLabel.Content = "Document ok.";
                statusLabel.Background = Brushes.Green;
            }
            catch (Exception e2)
            {
                statusLabel.Content = e2.Message;
                statusLabel.Background = Brushes.Red;
            }
        }

        private void TreeFocused(object sender, RoutedEventArgs e)
        {
            ReloadTreeView();
        }

        private void ReloadTreeView()
        {
            try
            {
                dataProvider.Document = Report.ConvertedXml;
                statusLabel.Content = "Document ok.";
                statusLabel.Background = Brushes.Green;
            }
            catch (Exception e2)
            {
                statusLabel.Content = e2.Message;
                statusLabel.Background = Brushes.Red;
            }
        }

    }
}
