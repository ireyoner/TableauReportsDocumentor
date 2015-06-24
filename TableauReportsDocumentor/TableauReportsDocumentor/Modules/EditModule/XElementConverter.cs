/*
 * File: XElementConverter.cs
 * Class: XElementConverter
 * 
 * Class responsible for delivering data to build report tree
 * 
 */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Xml;

namespace TableauReportsDocumentor.Modules.EditModule
{
    public class XElementConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            XmlElement element = value as XmlElement;
            if (element == null) return null;
            return element.SelectNodes("sections/section" +
                                       "|subsections/subsection" +
                                       "|content/text" +
                                       "|content/table" +
                                       "|header" +
                                       "|cell" +
                                       "|hcell" +
                                       "|rows" +
                                       "|row");
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
