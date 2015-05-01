using ICSharpCode.AvalonEdit.Folding;
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
using TableauReportsDocumentor.Data;

namespace TableauReportsDocumentor.Modules.EditModule
{
    /// <summary>
    /// Interaction logic for ReportXmlView.xaml
    /// </summary>
    public partial class ReportXmlView : UserControl
    {
        private FoldingManager foldingManager;
        private XmlFoldingStrategy foldingStrategy;
        private ReportContent report;
        public ReportContent Report
        {
            get { return report; }
            set
            {
                report = value;
                ReloadEditorView();
            }
        }

        public Label statusLabel { get; set; }

        private void ReloadEditorView()
        {
            if (Report != null)
                outputTest.Text = Report.ConvertedXmlAsString;
        }

        public ReportXmlView()
        {
            InitializeComponent();
            foldingManager = FoldingManager.Install(outputTest.TextArea);
            foldingStrategy = new XmlFoldingStrategy();
        }

        private void EditorRapaint(object sender, EventArgs e)
        {
            foldingStrategy.UpdateFoldings(foldingManager, outputTest.Document);
            try
            {
                Report.ConvertedXmlAsString = outputTest.Text;
                statusLabel.Content = "Document ok.";
                statusLabel.Background = Brushes.Green;
            }
            catch (Exception e2)
            {
                statusLabel.Content = e2.Message;
                statusLabel.Background = Brushes.Red;
            }
        }

        private void EditorFocused(object sender, RoutedEventArgs e)
        {
            ReloadEditorView();
        }

    }
}
