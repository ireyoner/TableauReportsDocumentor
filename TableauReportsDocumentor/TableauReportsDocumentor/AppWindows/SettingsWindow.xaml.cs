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
using TableauReportsDocumentor.Properties;

namespace TableauReportsDocumentor.AppWindows
{
    /// <summary>
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
        }

        string GetRelativePath(string filespec)
        {
            string folder = Directory.GetCurrentDirectory();
            Uri pathUri = new Uri(filespec);

            // Folders must end in a slash
            if (!folder.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                folder += Path.DirectorySeparatorChar;
            }
            Uri folderUri = new Uri(folder);
            return Uri.UnescapeDataString(folderUri.MakeRelativeUri(pathUri).ToString().Replace('/', Path.DirectorySeparatorChar));
        }

        private void SaveSettings(object sender, RoutedEventArgs e)
        {
            Settings.Default.UserConverter = UserConverterTB.Text;
            Settings.Default.ImportConvertersLocalization = ImportConvertersLocalizationTB.Text;
            Settings.Default.ImportConverterDefaultInstance = ImportConverterDefaultInstanceTB.Text;
            Settings.Default.ReportDocumentValidator = ReportDocumentValidatorTB.Text;
            Settings.Default.ImportConvertersAutoSearch = ImportConvertersAutoSearchRB.IsChecked??true;
            Settings.Default.Save();
            this.Close();
        }

        private void ImportConvertersLocalizationChange(object sender, RoutedEventArgs e)
        {
            using (System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog())
            {
                dlg.SelectedPath = Path.GetFullPath(ImportConverterDefaultInstanceTB.Text);
                dlg.ShowNewFolderButton = false;
                var result = dlg.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    ImportConvertersLocalizationTB.Text = GetRelativePath(dlg.SelectedPath);
                }
            }
        }

        private void ImportConverterDefaultInstanceChange(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "xsl file|*.xsl";
            if (openFileDialog.ShowDialog() == true)
            {
                ImportConverterDefaultInstanceTB.Text = GetRelativePath(Path.GetFullPath(openFileDialog.FileName));
            }
        }

        private void UserConverterChange(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "xsl file|*.xsl";
            if (openFileDialog.ShowDialog() == true)
            {
                UserConverterTB.Text = GetRelativePath(Path.GetFullPath(openFileDialog.FileName));
            }
        }

        private void ReportDocumentValidatorChange(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "xsd file|*.xsd";
            if (openFileDialog.ShowDialog() == true)
            {
                ReportDocumentValidatorTB.Text = GetRelativePath(Path.GetFullPath(openFileDialog.FileName));
            }
        }

        private void CancelSettings(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }

    [ValueConversion(typeof(bool), typeof(bool))]
    public class InverseBooleanConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        #endregion
    }
}
