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
    class DocxExport : ExportInterface
    {
        public string menuItemText
        {
            get { return "docx"; }
        }

        public BitmapFrame menuItemIcone
        {
            get
            {
                return null;
                //Uri iconUri = new Uri("docx.ico", UriKind.RelativeOrAbsolute);
                //return BitmapFrame.Create(iconUri);
            }
        }

        public object toolBarButtonContent
        {
            get { return menuItemText; }
        }

        public string exportSaveFileDialogFilter
        {
            get { return "Docx document|*.docx"; }
        }

        public bool export(string exportFileName, System.Xml.XmlDocument exportSource)
        {


            XmlNode root = exportSource.DocumentElement;
            string text;
            XmlNode nodeTest;
            using (DocX document = DocX.Create(exportFileName))
            {
                Paragraph p = document.InsertParagraph("Docx Test Document");

                foreach (XmlNode section in root.SelectNodes("//section"))
                {
                    nodeTest =  section.SelectSingleNode("text");
                    if ( nodeTest == null ){
                        text = string.Format("{0}", section.SelectSingleNode("title").InnerText);
                    }else{
                        text = string.Format("{0} -- {1}", section.SelectSingleNode("title").InnerText, nodeTest.InnerText);
                    }
                    p = document.InsertParagraph();
                    p.AppendLine(text);

                    foreach (XmlNode subsection in section.SelectNodes("subsection"))
                    {

                        p = document.InsertParagraph(); 

                        nodeTest = subsection.SelectSingleNode("text");
                        if (nodeTest == null)
                        {
                            text = string.Format("{0}", subsection.SelectSingleNode("title").InnerText);
                        }
                        else
                        {
                            text = string.Format("{0} -- {1}", subsection.SelectSingleNode("title").InnerText, nodeTest.InnerText);
                        }
                        p.AppendLine(text);

                        foreach (XmlNode table in subsection.SelectNodes("table"))
                        {
                            p = document.InsertParagraph(); 
                            p.AppendLine(table.SelectSingleNode("title").InnerText);

                            XmlNodeList headers = table.SelectNodes("header/cell");
                            int colCount = headers.Count;
                            Table t =document.InsertTable(1, colCount);
                            t.Design = TableDesign.ColorfulListAccent1;
                            int i;
                            for (i = 0; i < colCount; i++ )
                            {
                                t.Rows[0].Cells[i].InsertParagraph(headers[i].InnerText);
                            }

                            XmlNodeList rows = table.SelectNodes("row");
                            foreach (XmlNode row in rows)
                            {
                                Row newTabRow = t.InsertRow();
                                 for (i = 0; i < colCount; i++ )
                                 {
                                    newTabRow.Cells[i].InsertParagraph(row.ChildNodes[i].InnerText);
                                 }

                            }

                        }

                    }

                    foreach (XmlNode table in section.SelectNodes("table"))
                    {
                        p = document.InsertParagraph();
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

                        XmlNodeList rows = table.SelectNodes("row");
                        foreach (XmlNode row in rows)
                        {
                            Row newTabRow = t.InsertRow();
                            for (i = 0; i < colCount; i++)
                            {
                                newTabRow.Cells[i].InsertParagraph(row.ChildNodes[i].InnerText);
                            }

                        }

                    }
                }
                document.Save();
            }

            //throw new NotImplementedException();
            return true;
        }

    }
}
