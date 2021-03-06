﻿/*
 * File: ExportCSV.cs
 * Class: ExportCSV
 * 
 * Exporter class implementing ExportInterface and ExporterInstance that transforms 
 * our TRD xml file into a complete .csv file.
 * 
 * Note: (!!!)
 * Functions undocumented here that override initial implementation from ExporterInstance 
 * are documented in ExporterInstance source where they were first declared.
 * 
 * It uses only native C# functions.
 * 
 */

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
    class ExportCSV : ExporterInstance, ExportInterface
    {

        private String fieldsSeparator = ";";

        /*
         * Property: FileExtension
         * Read only property stating the extension of the file produced by this exporter.
         */
        public string FileExtension
        {
            get { return "csv"; }
        }

        /*
         * Property: MenuItemText
         * Read only property providing the name of the menu item under which this exporter will be avaliable.
         */
        public string MenuItemText
        {
            get { return "csv"; }
        }

        public BitmapFrame MenuItemIcone
        {
            get
            {
                return null;
                //Uri iconUri = new Uri("csv.ico", UriKind.RelativeOrAbsolute);
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
            get { return "csv file|*.csv"; }
        }

        /*
         * Function: StringToCSVCell
         * Converts a string into a single cell of a CSV file
         * 
         * Parameters: 
         *  str - a text string we want to convert
         *  
         * Returns:
         *  A string containing the input string after clearing of any field separators with a field separator added at the end
         *  
         */
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

        /*
         * Function: CreateSubSection
         * Creates a subsection in a csv file
         * 
         * Parameters:
         *  csvBuilder - a StringBuilder object holding a recipe for building a csv file at the current stage
         *  subSection - a TRD xml node containing subsection data to be exported
         *  
         * Returns:
         *  a StringBuilder object expanded with creation of the input subsection in csv form.
         */
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

        /*
         * Function: CreateText
         * A function creating a cell with given text in the csv file
         * 
         * Parameters:
         *  csvBuilder - a StringBuilder object holding a recipe for building a csv file at the current stage
         *  text - a TRD xml node containing text to be exported
         *  
         * Returns:
         *  a StringBuilder object expanded with creation of a cell with given text
         */
        private StringBuilder CreateText(StringBuilder csvBuilder, XmlNode text)
        {
            csvBuilder.AppendLine(StringToCSVCell(text.InnerText));
            return csvBuilder;
        }

        /*
         * Function: CreateTable
         * A function creating a table with given contents in the csv form
         * 
         * Parameters:
         *  csvBuilder - a StringBuilder object holding a recipe for building a csv file at the current stage
         *  table - a TRD xml node containing all the table data to be exported
         *  
         * Returns:
         *  a StringBuilder object expanded with creation of the given table in a csv form
         */
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

        StringBuilder csvBuilder;
        /* Function: ExportInit
         * Initialisating export of a csv file.
         * 
         * Parameters:
         *  reportTitle - title of the report
         *  rootNode - root node of the TRD xml file
         *  exportFileName - path to the desired output file
         *  
         */
        protected override void ExportInit(string reportTitle, XmlNode rootNode, string exportFileName)
        {
            csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("Report;" + StringToCSVCell(reportTitle));
        }

        /*
         * Function: ExportEnd
         * Finish the process of exporting to a csv file and save it.
         * 
         * Parameters:
         *  reportTitle - title of the report
         *  rootNode - root node of the TRD xml file
         *  exportFileName - path to the desired output file
         *  
         */
        protected override bool ExportEnd(string reportTitle, XmlNode rootNode, string exportFileName)
        {
            File.WriteAllText(exportFileName, csvBuilder.ToString());
            return true;
        }

        protected override void SectionInit(string sectionTitle, XmlNode sectionNode)
        {
            csvBuilder.AppendLine();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("Section;" + StringToCSVCell(sectionTitle));
        }

        protected override void SectionContentStart(string sectionTitle, XmlNode sectionNode, int contentNodesCount) { }
        protected override void SectionContentEnd(string sectionTitle, XmlNode sectionNode) { }
        protected override void SectionSubsectionsStart(string sectionTitle, XmlNode sectionNode, int subsectionNodesCount) { }
        protected override void SectionSubsectionsEnd(string sectionTitle, XmlNode sectionNode) { }
        protected override void SectionEnd(string sectionTitle, XmlNode sectionNode) { }

        protected override void SubSectionInit(string subSectionTitle, XmlNode subSectionNode)
        {
            csvBuilder.AppendLine();
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("SubSection;" + StringToCSVCell(subSectionTitle));
        }

        protected override void SubSectionContentStart(string subSectionTitle, XmlNode subSectionNode, int contentNodesCount) { }
        protected override void SubSectionContentEnd(string subSectionTitle, XmlNode subSectionNode) { }
        protected override void SubSectionEnd(string subSectionTitle, XmlNode subSectionNode) { }

        protected override void CreateText(XmlNode textNode)
        {
            csvBuilder.AppendLine("Text;" + StringToCSVCell(textNode.InnerText));
        }

        protected override void TableInit(string tableTitle, XmlNode tableNode)
        {
            csvBuilder.AppendLine();
            csvBuilder.AppendLine("Table;" + StringToCSVCell(tableTitle));
        }

        protected override void TableHeaderStart(string tableTitle, XmlNode headerNode, int headerNodesCount) { }

        protected override void TableHeaderCell(string tableTitle, XmlNode cellNode)
        {
            csvBuilder.Append(StringToCSVCell(cellNode.InnerText));
        }

        protected override void TableHeaderEnd(string tableTitle, XmlNode headerNode)
        {
            csvBuilder.AppendLine();
        }

        protected override void TableRowsStart(string tableTitle, XmlNode rowsNode, int contentNodesCount) { }
        protected override void TableRowStart(string tableTitle, XmlNode rowNode, int contentNodesCount) { }
        protected override void TableRowCell(string tableTitle, XmlNode cellNode)
        {
           csvBuilder.Append(StringToCSVCell(cellNode.InnerText));
        }

        protected override void TableRowEnd(string tableTitle, XmlNode rowNode)
        {
            csvBuilder.AppendLine();
        }

        protected override void TableRowsEnd(string tableTitle, XmlNode rowsNode) { }
        protected override void TableEnd(string tableTitle, XmlNode tableNode) { }
    }
}
