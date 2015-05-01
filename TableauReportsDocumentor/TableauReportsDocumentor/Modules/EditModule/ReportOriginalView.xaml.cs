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
    /// Interaction logic for ReportOriginalSource.xaml
    /// </summary>
    public partial class ReportOriginalView : UserControl
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

        private void ReloadEditorView()
        {
            outputTest.Text = Report.Original;
            foldingStrategy.UpdateFoldings(foldingManager, outputTest.Document);
        }

        public ReportOriginalView()
        {
            InitializeComponent();
            foldingManager = FoldingManager.Install(outputTest.TextArea);
            foldingStrategy = new XmlFoldingStrategy();
        }

    }
}
