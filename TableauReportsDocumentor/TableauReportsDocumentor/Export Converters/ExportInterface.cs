using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Xml;

namespace TableauReportsDocumentor.Export_Converters
{
    interface ExportInterface
    {
        // Required field, visible in Import menu. If this field is null, then ExportInterface is not registered
        String menuItemText { get; }

        // Optional icone visible in Import menu next to menuItemText
        BitmapFrame menuItemIcone { get; }

        // Optional field, if not null Button with it is created in Export ToolBar
        Object toolBarButtonContent { get; }

        // Optional, if not null, the SaveFileDialog with Filter = exportSaveFileDialogFilter is shown when exporting,
        //           else, if null, SaveFileDialog is not displayed
        String exportSaveFileDialogFilter { get; }

        // Method called when button or menu item is clicked
        bool export(String exportFileName, XmlDocument exportSource);
    }
}
