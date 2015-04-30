using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media.Imaging;
using System.Xml;
using TableauReportsDocumentor.Data;

namespace TableauReportsDocumentor.Modules.ImportModule
{
    public interface ImportInterface
    {
        // Required field, visible in Import menu. If this field is null, then ExportInterface is not registered
        String MenuItemText { get; }

        // Optional icone visible in Import menu next to menuItemText
        BitmapFrame MenuItemIcone { get; }

        // Method called when button or menu item is clicked
        ImportedDocument Import();

    }
}
