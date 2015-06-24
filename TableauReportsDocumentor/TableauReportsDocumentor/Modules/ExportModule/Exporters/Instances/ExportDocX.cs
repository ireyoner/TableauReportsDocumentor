/*
 * File: ExportDocX.cs
 * Class: ExportDocX
 * 
 * Exporter class implementing ExportInterface that transforms our TRD xml file into a complete .docx document.
 * It uses a DocX library to accomplish it's task.
 * 
 * Library: DocX
 * Version: 1.0.0.15
 * Source: https://docx.codeplex.com/
 */
 
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
        /*
         * Property: FileExtension
         * Read only property stating the extension of the file produced by this exporter.
         */
        public string FileExtinsion
        {
            get { return "docx"; }
        }

        /*
         * Property: MenuItemText
         * Read only property providing the name of the menu item under which this exporter will be avaliable.
         */
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

        /*
         * Property: ExportSaveFileDialogFilter
         * Read only property providing filter name and value for the SaveFile Dialog.
         */
        public string ExportSaveFileDialogFilter
        {
            get { return "Docx document|*.docx"; }
        }

        /*
         * Function Export
         * Main function of the class. Performs the export process of our TRD xml into a .docx file
         * 
         * Parameters:
         *  exportFileName - string holding a path to the target location for the output file
         *  exportSource - XML Document object holding TRD xml to be exported
         *  
         * Returns:
         *  A bool value stating successfull export
         *  
         */
        public bool Export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            XmlNode root = exportSource.DocumentElement;
            // Creation of the document object
            DocX document = DocX.Create(exportFileName);
            // Insert the name of the exported workbook at the start of the document
            Paragraph p = document.InsertParagraph();
            p.AppendLine(root.SelectSingleNode("title").InnerText).Heading(HeadingType.Heading1).Color(System.Drawing.Color.Black);
            // Create each section of the document
            foreach (XmlNode section in root.SelectNodes("//section"))
            {
                document = CreateSection(document, section);
            }
            // Saving of the document
            document.Save();

            return true;
        }

        /*
         * Function: CreateSection
         * Adds a new section to the created output document based on the passed TRD xml of said section.
         * 
         * Parameters:
         *  document - current document object
         *  section - node from the TRD xml holding the data of the new section that is going to be added.
         * 
         * Returns:
         *  document object with a new section corresponding to provided TRD xml appended at the end.
         *
         */

        private DocX CreateSection(DocX document, XmlNode section)
        {
            // Insert a header with a section title.
            var p = document.InsertParagraph(section.SelectSingleNode("title").InnerText).Heading(HeadingType.Heading2).Color(System.Drawing.Color.Black);
           
            // Insert contents of a the currently exported section if they are avaliable
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

            // Create any subsections present in the currently exported section.
            foreach (XmlNode subsection in section.SelectNodes(".//subsection"))
            {
                document = CreateSubSection(document, subsection);
            }
            return document;
        }

        /*
         * Function: CreateSubSection
         * Adds a new subsection to the created output document based on the passed TRD xml of said subsection.
         * 
         * Parameters:
         *  document - current document object
         *  subSection - node from the TRD xml holding the data of the new subsection that is going to be added.
         * 
         * Returns:
         *  document object with a new subsection corresponding to provided TRD xml appended at the end.
         *
         */
        private DocX CreateSubSection(DocX document, XmlNode subSection)
        {
            //Insert title of the subsection
            var p = document.InsertParagraph(subSection.SelectSingleNode("title").InnerText).Heading(HeadingType.Heading4).Color(System.Drawing.Color.Black);

            //Insert subsection contents
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

        /*
         * Function: CreateText
         * Inserts text from the provided XML node into the document.
         * 
         * Parameters:
         *  document - current document object
         *  text - a XML node containing text to be added
         *  
         * Returns:
         *  document object with text corresponding to the provided XML node appended at the end
         *
         */
        private DocX CreateText(DocX document, XmlNode text)
        {
            document.Paragraphs.Last().AppendLine(text.InnerText).FontSize(10).Color(System.Drawing.Color.Black);
            return document;
        }

        /*
         * Function: CreateTable
         * Creates a new table in the document based on provided TRD xml data.
         * 
         * Parameters:
         *  document - current document object
         *  table - a XML node containing the data of the table that is going to be added
         *  
         * Returns: 
         *  document object with the table corresponding to given data appended at the end.
         *
         */
        private DocX CreateTable(DocX document, XmlNode table)
        {
            //Insert a heading with the table name
            var p = document.InsertParagraph();
            p.AppendLine(table.SelectSingleNode("title").InnerText).Heading(HeadingType.Heading5).Color(System.Drawing.Color.Black);

            // Create table and its header row.
            XmlNodeList headers = table.SelectNodes("header/hcell");
            int colCount = headers.Count; 
            Table t = document.InsertTable(1, colCount);
            // Set all borders for the table cells to be drawn.
            setAllTableBorders(t, new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, System.Drawing.Color.Black));
            int i;
            for (i = 0; i < colCount; i++)
            {
                t.Rows[0].Cells[i].InsertParagraph(headers[i].InnerText);
            }

            // For each row described in the xml add a table row in the document
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

        /*
         * Function: setAllTableBorders
         * A helper function setting a drawing style of all table borders.
         * 
         * Parameters:
         *  tab - table that the style is to be set for
         *  border - an object of type Border with appropriate border style set
         *
         */
        private void setAllTableBorders(Table tab, Border border)
        {
            foreach (TableBorderType type in Enum.GetValues(typeof(TableBorderType)))
            {
                tab.SetBorder(type, border);
            }
        }
        

    }
}
