using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows.Media.Imaging;
using System.IO;

namespace TableauReportsDocumentor.Export_Converters
{
    class ExportCSV : ExportInterface
    {
        private String fieldsSeparator = ";";
        public string FileExtinsion
        {
            get { return "csv"; }
        }

        public string MenuItemText
        {
            get { return "csv"; }
        }

        public BitmapFrame MenuItemIcone
        {
            get
            {
                return null;
                //Uri iconUri = new Uri("docx.ico", UriKind.RelativeOrAbsolute);
                //return BitmapFrame.Create(iconUri);
            }
        }

        public object ToolBarButtonContent
        {
            get { return MenuItemText; }
        }

        public string ExportSaveFileDialogFilter
        {
            get { return "csv file|*.csv"; }
        }

        private string StringToCSVCell(string str)
        {
            if (str.Contains(fieldsSeparator) || str.Contains("\"") || str.Contains(System.Environment.NewLine))
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("\"");
                foreach (char nextChar in str)
                {
                    sb.Append(nextChar);
                    if (nextChar == '"')
                        sb.Append("\"");
                }
                sb.Append("\"");
                return sb.ToString() + fieldsSeparator;
            }
            return str + fieldsSeparator;
        }

        public bool Export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            XmlNode root = exportSource.DocumentElement;
            var csvBuilder = new StringBuilder();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("Section;" + StringToCSVCell(root.SelectSingleNode("title").InnerText));

            foreach (XmlNode section in root.SelectNodes("//section"))
            {
                csvBuilder = CreateSection(csvBuilder, section);
            }
            File.WriteAllText(exportFileName, csvBuilder.ToString()); 
            
            return true;
        }

        private StringBuilder CreateSection(StringBuilder csvBuilder, XmlNode section)
        {
            csvBuilder.AppendLine();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("Section;" + StringToCSVCell(section.SelectSingleNode("title").InnerText));

            foreach (XmlNode item in section.SelectNodes("table|text"))
            {
                if (item.Name == "text")
                {
                    csvBuilder = CreateText(csvBuilder, item);
                }
                else if (item.Name == "table")
                {
                    csvBuilder = CreateTable(csvBuilder, item);
                }
            }

            foreach (XmlNode subsection in section.SelectNodes("subsection"))
            {
                csvBuilder = CreateSubSection(csvBuilder, subsection);
            }
            return csvBuilder;
        }

        private StringBuilder CreateSubSection(StringBuilder csvBuilder, XmlNode subSection)
        {
            csvBuilder.AppendLine();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("SubSection;" + StringToCSVCell(subSection.SelectSingleNode("title").InnerText));

            foreach (XmlNode item in subSection.SelectNodes("table|text"))
            {
                if (item.Name == "text")
                {
                    csvBuilder = CreateText(csvBuilder, item);
                }
                else if (item.Name == "table")
                {
                    csvBuilder = CreateTable(csvBuilder, item);
                }
            }
            return csvBuilder;
        }

        private StringBuilder CreateText(StringBuilder csvBuilder, XmlNode text)
        {
            csvBuilder.AppendLine(StringToCSVCell(text.InnerText));
            return csvBuilder;
        }

        private StringBuilder CreateTable(StringBuilder csvBuilder, XmlNode table)
        {
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("Table;" + StringToCSVCell(table.SelectSingleNode("title").InnerText));

            XmlNodeList headers = table.SelectNodes("header/cell");
            int colCount = headers.Count;
            for (int i = 0; i < colCount; i++)
            {
                csvBuilder.Append(StringToCSVCell(headers[i].InnerText));
            }
            csvBuilder.AppendLine();

            XmlNodeList rows = table.SelectNodes("rows/row");
            foreach (XmlNode row in rows)
            {
                for (int i = 0; i < colCount; i++)
                {
                    csvBuilder.Append(StringToCSVCell(row.ChildNodes[i].InnerText));
                }
                csvBuilder.AppendLine();
            }
            return csvBuilder;
        }

    }
}
