/*
 * File: ExportInstance.cs
 * Class: ExportInstance
 * 
 * An abstract class providing some of the generic methods that can be used by multiple exporters.
 * Serves as a way to simplyfy and unify implementation of new exporters.
 * 
 * Important: Does not have to be always used when implementing an exporter class.
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TableauReportsDocumentor.Export_Converters
{
    abstract class ExporterInstance
    {
        /*
         * Function Export
         * Main function of the class. Performs the export process of our TRD xml to an output file
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
            var title = root.SelectSingleNode("title").InnerText;
            
            ExportInit(title, root, exportFileName);

            foreach (XmlNode section in root.SelectNodes("sections/section"))
            {
                CreateSection(section);
            }

            return ExportEnd(title, root, exportFileName);
        }

        /*
         * Function: ExportInit
         * Abstract function. Used by specific exporters to initialise the export process.
         * 
         * Parameters:
         *  reportTitle - title of the report
         *  rootNode - root node of the TRD xml file
         *  exportFileName - path to the desired output file
         *  
         */
        protected abstract void ExportInit(string reportTitle, XmlNode rootNode, string exportFileName);
        /*
         * Function: ExportEnd
         * Abstract function. Used by specific exporters to finish the export process.
         * 
         * Parameters:
         *  reportTitle - title of the report
         *  rootNode - root node of the TRD xml file
         *  exportFileName - path to the desired output file
         *  
         */
        protected abstract bool ExportEnd(string reportTitle, XmlNode rootNode, string exportFileName);

        /*
         * Function: CreateSection
         * Adds a new section to the exported file based on the passed TRD xml of said section.
         * Uses a number of abstract functions to perform this task. These functions need to be implemented in the specific exporters.
         * 
         * Parameters:
         *  section - node from the TRD xml holding the data of the new section that is going to be added.
         *
         */
        private void CreateSection(XmlNode section)
        {
            var title = section.SelectSingleNode("title").InnerText;

            SectionInit(title, section);

            SectionContentStart(title, section, section.SelectNodes("content/*").Count);
            foreach (XmlNode item in section.SelectNodes("content/*"))
            {
                if (item.Name == "text")
                {
                    CreateText(item);
                }
                else if (item.Name == "table")
                {
                    CreateTable(item);
                }
            }
            SectionContentEnd(title, section);

            SectionSubsectionsStart(title, section, section.SelectNodes("subsections/subsection").Count);
            foreach (XmlNode subsection in section.SelectNodes("subsections/subsection"))
            {
                CreateSubSection(subsection);
            }
            SectionSubsectionsEnd(title, section);

            SectionEnd(title, section);
        }

        /*
         * Function: SectionInit
         * Abstract function. Used by specific exporters to initialise exporting a section.
         * 
         * Parameters:
         *  sectionTitle - title of the section
         *  sectionNode - a node of the TRD xml file holding the section data
         *  
         */
        protected abstract void SectionInit(string sectionTitle, XmlNode sectionNode);
        /*
         * Function: SectionContentStart
         * Abstract function. Used by specific exporters to perform any operations needed at the start of exporting section contents.
         * 
         * Parameters:
         *  sectionTitle - title of the section
         *  sectionNode - a node of the TRD xml file holding the section data
         *  contentNodesCount - number of nodes in the section
         */
        protected abstract void SectionContentStart(string sectionTitle, XmlNode sectionNode, int contentNodesCount);

        /*
         * Function: SectionSubsectionsStart
         * Abstract function. Used by specific exporters to perform any operations needed before exporting any subsections of a section 
         * are exported.
         * 
         * Parameters:
         *  sectionTitle - title of the section
         *  sectionNode - a node of the TRD xml file holding the section data
         *  subsectionNodesCount - number of nodes holding section subsections
         */
        protected abstract void SectionSubsectionsStart(string sectionTitle, XmlNode sectionNode, int subsectionNodesCount);
        /*
         * Function: SectionSubsectionsStart
         * Abstract function. Used by specific exporters to perform any operations needed after exporting all subsections of a section 
         * are exported.
         * 
         * Parameters:
         *  sectionTitle - title of the section
         *  sectionNode - a node of the TRD xml file holding the section data
         */
        protected abstract void SectionSubsectionsEnd(string sectionTitle, XmlNode sectionNode);

        /*
         * Function: SectionContentEnd
         * Abstract function. Used by specific exporters to perform any operations needed after exporting section contents
         * 
         * Parameters:
         *  sectionTitle - title of the section
         *  sectionNode - a node of the TRD xml file holding the section data
         */
        protected abstract void SectionContentEnd(string sectionTitle, XmlNode sectionNode);

        /*
         * Function: SectionEnd
         * Abstract function. Used by specific exporters to finish exporting a section.
         * 
         * Parameters:
         *  sectionTitle - title of the section
         *  sectionNode - a node of the TRD xml file holding the section data
         *  
         */
        protected abstract void SectionEnd(string sectionTitle, XmlNode sectionNode);

        /*
         * Function: CreateSubSection
         * Adds a new subsection to the exported file based on the passed TRD xml of said subsection.
         * Uses a number of abstract functions to perform this task. These functions need to be implemented in the specific exporters.
         * 
         * Parameters:
         *  subSection - node from the TRD xml holding the data of the new subsection that is going to be added.
         *
         */
        private void CreateSubSection(XmlNode subSection)
        {
            var title = subSection.SelectSingleNode("title").InnerText;

            SubSectionInit(title, subSection);

            SubSectionContentStart(title, subSection, subSection.SelectNodes("content/*").Count);
            foreach (XmlNode item in subSection.SelectNodes("content/*"))
            {
                if (item.Name == "text")
                {
                    CreateText(item);
                }
                else if (item.Name == "table")
                {
                    CreateTable(item);
                }
            }
            SubSectionContentEnd(title, subSection);

            SubSectionEnd(title, subSection);
        }

        /*
         * Function: SubSectionInit
         * Abstract function. Used by specific exporters to initialise exporting a subsection.
         * 
         * Parameters:
         *  subSectionTitle - title of the subsection
         *  subSectionNode - a node of the TRD xml file holding the subsection data
         *  
         */
        protected abstract void SubSectionInit(string subSectionTitle, XmlNode subSectionNode);
        /*
         * Function: SubSectionContentStart
         * Abstract function. Used by specific exporters to perform any operations needed at the start of exporting subsection contents.
         * 
         * Parameters:
         *  subSectionTitle - title of the subsection
         *  subSectionNode - a node of the TRD xml file holding the subsection data
         *  contentNodesCount - number of nodes in the subsection
         */
        protected abstract void SubSectionContentStart(string subSectionTitle, XmlNode subSectionNode, int contentNodesCount);
        /*
         * Function: SubSectionContentEnd
         * Abstract function. Used by specific exporters to perform any operations needed after exporting subsection contents
         * 
         * Parameters:
         *  subSectionTitle - title of the subsection
         *  subSectionNode - a node of the TRD xml file holding the subsection data
         */
        protected abstract void SubSectionContentEnd(string subSectionTitle, XmlNode subSectionNode);
        /*
         * Function: SubSectionEnd
         * Abstract function. Used by specific exporters to finish exporting a subsection.
         * 
         * Parameters:
         *  subSectionTitle - title of the subsection
         *  subSectionNode - a node of the TRD xml file holding the subsection data
         *  
         */
        protected abstract void SubSectionEnd(string subSectionTitle, XmlNode subSectionNode);

        /*
         * Function: CreateText
         * Abstract function. Used by specific exporters to export text from text nodes of TRD file
         * 
         * Parameters:
         *  textNode - a node of the TRD xml file holding the text
         */
        protected abstract void CreateText(XmlNode textNode);

        /*
         * Function: CreateTable
         * Adds a new table to the exported file based on the passed TRD xml of said table.
         * Uses a number of abstract functions to perform this task. These functions need to be implemented in the specific exporters.
         * 
         * Parameters:
         *  table - node from the TRD xml holding the data of the new table that is going to be added.
         *
         */
        private void CreateTable(XmlNode table)
        {
            var title = table.SelectSingleNode("title").InnerText;

            TableInit(title, table);

            var header = table.SelectSingleNode("header");
            var headerCells = header.SelectNodes("hcell");
            int colCount = headerCells.Count;

            TableHeaderStart(title, header, colCount);
            foreach (XmlNode cell in header.SelectNodes("hcell"))
            {
                TableHeaderCell(title, cell);
            }
            TableHeaderEnd(title, header);

            var rows = table.SelectSingleNode("rows");
            var rowsRows = rows.SelectNodes("row");
            int rowsCount = rowsRows.Count;

            TableRowsStart(title, rows, rowsCount);
            foreach (XmlNode row in rowsRows)
            {
                var rowCells = row.SelectNodes("cell");
                int cellsCount = rowCells.Count;

                TableRowStart(title, row, cellsCount);
                foreach (XmlNode cell in rowCells)
                {
                    TableRowCell(title, cell);
                }
                TableRowEnd(title, row);
            }
            TableHeaderEnd(title, header);

            TableEnd(title, header);
        }

        /*
         * Function: TableInit
         * Abstract function. Used by specific exporters to initialise exporting a table.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  tableNode - TRD xml node holding the table data
         */
        protected abstract void TableInit(string tableTitle, XmlNode tableNode);
        /*
         * Function: TableHeaderStart
         * Abstract function. Used by specific exporters to perform any operations needed before exporting 
         * specific cells creating the header of a table.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  headerNode - TRD xml node holding the headers of a table
         *  headerNodesCount - number of header nodes - number of table columns
         */
        protected abstract void TableHeaderStart(string tableTitle, XmlNode headerNode, int headerNodesCount);
        /*
         * Function: TableHeaderCell
         * Abstract function. Used by specific exporters to export a cell of a table header.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  cellNode - TRD xml node holding a single header cell data
         */
        protected abstract void TableHeaderCell(string tableTitle, XmlNode cellNode);
        /*
         * Function: TableHeaderEnd
         * Abstract function. Used by specific exporters to perorm any operations needed after exporting
         * a table header.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  headerNode - TRD xml node holding the headers of a table
         */
        protected abstract void TableHeaderEnd(string tableTitle, XmlNode headerNode);
        /*
         * Function: TableRowsStart
         * Abstract function. Used by specific exporters to perform any operations needed before exporting 
         * any of the table rows.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  rowsNode - TRD xml node holding all the rows of a table
         *  contentNodesCount - number of header nodes - number of table rows
         */
        protected abstract void TableRowsStart(string tableTitle, XmlNode rowsNode, int contentNodesCount);
        /*
         * Function: TableRowStart
         * Abstract function. Used by specific exporters to perform any operations needed before exporting 
         * specific cells of a table row.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  rowNode - TRD xml node holding data of a single row
         *  contentNodesCount - number of header nodes - number of table columns
         */
        protected abstract void TableRowStart(string tableTitle, XmlNode rowNode, int contentNodesCount);
        /*
         * Function: TableRowCell
         * Abstract function. Used by specific exporters export a single cell of a table.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  cellNode - TRD xml node holding the table cell node
         *
         */
        protected abstract void TableRowCell(string tableTitle, XmlNode cellNode);
        /*
         * Function: TableRowEnd
         * Abstract function. Used by specific exporters to perform any operations needed after exporting 
         * specific cells of a table row.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  rowNode - TRD xml node holding data of a single row
         */
        protected abstract void TableRowEnd(string tableTitle, XmlNode rowNode);
        /*
         * Function: TableRowsEnd
         * Abstract function. Used by specific exporters to perform any operations needed before after 
         * all of the table rows.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  rowsNode - TRD xml node holding all the rows of a table
         */
        protected abstract void TableRowsEnd(string tableTitle, XmlNode rowsNode);
        /*
         * Function: TableEnd
         * Abstract function. Used by specific exporters to finish exporting a table.
         * 
         * Parameters:
         *  tableTitle - title of the table
         *  tableNode - TRD xml node holding the table data
         */
        protected abstract void TableEnd(string tableTitle, XmlNode tableNode);

    }
}
