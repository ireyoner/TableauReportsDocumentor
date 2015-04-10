using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.Xml;
using System.Windows.Media.Imaging;

namespace TableauReportsDocumentor.Export_Converters
{
    class ExportDocX : ExportInterface
    {

        public string FileExtinsion
        {
            get { return "docx"; }
        }

        public string MenuItemText
        {
            get { return "docx"; }
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
            get { return "Docx document|*.docx"; }
        }

        public bool Export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            XmlNode root = exportSource.DocumentElement;
            DocX document = DocX.Create(exportFileName);
            Paragraph p = document.InsertParagraph();
            p.AppendLine(root.SelectSingleNode("title").InnerText);

            foreach (XmlNode section in root.SelectNodes("//section"))
            {
                document = CreateSection(document, section);
            }
            document.Save();

            return true;
        }

        private DocX CreateSection(DocX document, XmlNode section)
        {
            var p = document.InsertParagraph(section.SelectSingleNode("title").InnerText);

            foreach (XmlNode item in section.SelectNodes("content/*"))
            {
                if (item.Name == "text")
                {
                    document = CreateText(document, item);
                }
                else if (item.Name == "table")
                {
                    document = CreateTable(document, item);
                }
            }

            foreach (XmlNode subsection in section.SelectNodes(".//subsection"))
            {
                document = CreateSubSection(document, subsection);
            }
            return document;
        }

        private DocX CreateSubSection(DocX document, XmlNode subSection)
        {
            var p = document.InsertParagraph(subSection.SelectSingleNode("title").InnerText);

            foreach (XmlNode item in subSection.SelectNodes("content/*"))
            {
                if (item.Name == "text")
                {
                    document = CreateText(document, item);
                }
                else if (item.Name == "table")
                {
                    document = CreateTable(document, item);
                }
            }
            return document;
        }

        private DocX CreateText(DocX document, XmlNode text)
        {
            document.Paragraphs.Last().AppendLine(text.InnerText);
            return document;
        }

        private DocX CreateTable(DocX document, XmlNode table)
        {
            var p = document.InsertParagraph();
            p.AppendLine(table.SelectSingleNode("title").InnerText);

            XmlNodeList headers = table.SelectNodes("header/cell");
            int colCount = headers.Count;
            Table t = document.InsertTable(1, colCount);
            t.Design = TableDesign.ColorfulListAccent1;
            int i;
            for (i = 0; i < colCount; i++)
            {
                t.Rows[0].Cells[i].InsertParagraph(headers[i].InnerText);
            }

            XmlNodeList rows = table.SelectNodes("rows/row");
            foreach (XmlNode row in rows)
            {
                Row newTabRow = t.InsertRow();
                for (i = 0; i < colCount; i++)
                {
                    newTabRow.Cells[i].InsertParagraph(row.ChildNodes[i].InnerText);
                }

            }

            return document;
        }

    }
}
