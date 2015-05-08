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
using System.Windows.Shapes;

namespace TableauReportsDocumentor.Modules.ExportModule
{
    /// <summary>
    /// Interaction logic for ExportFileError.xaml
    /// </summary>
    public partial class ExportFileError : Window
    {
        public ExportFileError()
        {
            InitializeComponent();
        }

        public ExportFileError(String filepath)
        {
            InitializeComponent();
            this.ErrorMessage = filepath;
        }

        public String ErrorMessage
        {
            get { return ErrorMessageText.Text; }
            set { ErrorMessageText.Text = value; }
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
